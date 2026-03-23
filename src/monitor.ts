import type { OpenClawConfig, RuntimeEnv } from "openclaw/plugin-sdk";
import type { A365Config, A365MessageMetadata } from "./types.js";
import { getA365Runtime } from "./runtime.js";
import { runWithGraphToolContext } from "./graph-tools.js";
import { resolveA365Credentials } from "./token.js";
import { saveConversationReference } from "./conversation-store.js";
import {
  registerAdapter,
  resolveAccountIdByRecipientId,
  getAdapterByRecipientId,
  setAdapter,
  setBlueprintClientId,
} from "./adapter-store.js";
import { resolveAccountA365Config } from "./channel.js";

/** Guard against double-start: tracks whether the a365 server is already listening. */
let a365ServerActive = false;

export type MonitorA365Opts = {
  cfg: OpenClawConfig;
  runtime?: RuntimeEnv;
  abortSignal?: AbortSignal;
};

export type MonitorA365Result = {
  app: unknown;
  shutdown: () => Promise<void>;
};

/**
 * Activity shape for metadata extraction.
 */
export type ActivityForMetadata = {
  from?: { id: string; name?: string; aadObjectId?: string };
  recipient?: { id: string; name?: string };
  conversation?: { id: string; isGroup?: boolean; tenantId?: string };
  serviceUrl?: string;
  id?: string;
  channelId?: string;
  locale?: string;
  channelData?: {
    tenant?: { id: string };
    team?: { id: string; name?: string };
    channel?: { id: string; name?: string };
  };
};

/**
 * Extract message metadata from an Agents SDK activity.
 */
export function extractMessageMetadata(activity: ActivityForMetadata): A365MessageMetadata {
  return {
    userId: activity.from?.id || "",
    userEmail: activity.from?.aadObjectId || activity.from?.id,
    userName: activity.from?.name,
    userAadId: activity.from?.aadObjectId,
    conversationId: activity.conversation?.id || "",
    isGroup: activity.conversation?.isGroup || false,
    tenantId: activity.conversation?.tenantId || activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl || "",
    activityId: activity.id,
    channelId: activity.channelId,
    teamId: activity.channelData?.team?.id,
    teamName: activity.channelData?.team?.name,
    channelName: activity.channelData?.channel?.name,
  };
}

/**
 * Build a StoredConversationReference from an activity for proactive messaging.
 */
export function buildConversationReference(activity: ActivityForMetadata): StoredConversationReference {
  return {
    conversationId: activity.conversation?.id || "",
    serviceUrl: activity.serviceUrl || "",
    channelId: activity.channelId || "msteams",
    botId: activity.recipient?.id || "",
    botName: activity.recipient?.name,
    userId: activity.from?.id || "",
    userName: activity.from?.name,
    userAadId: activity.from?.aadObjectId,
    tenantId: activity.conversation?.tenantId || activity.channelData?.tenant?.id,
    isGroup: activity.conversation?.isGroup || false,
    locale: activity.locale,
    updatedAt: Date.now(),
  };
}

/**
 * Set the SDK environment variables for a specific account's credentials.
 * Must be called BEFORE creating the AgentApplication/CloudAdapter since
 * the SDK reads env vars at construction time and caches them.
 */
function setEnvForAccount(creds: { appId: string; appPassword: string; tenantId: string }, port: number): void {
  process.env["connections__serviceConnection__settings__clientId"] = creds.appId;
  process.env["connections__serviceConnection__settings__clientSecret"] = creds.appPassword;
  process.env["connections__serviceConnection__settings__tenantId"] = creds.tenantId;
  process.env["connectionsMap__0__connection"] = "serviceConnection";
  process.env["connectionsMap__0__serviceUrl"] = "*";
  process.env.MicrosoftAppId = creds.appId;
  process.env.MicrosoftAppPassword = creds.appPassword;
  process.env.MicrosoftAppTenantId = creds.tenantId;
  process.env.MicrosoftAppType = "SingleTenant";
  process.env.PORT = String(port);
}

/**
 * Register the message handler on an AgentApplication instance.
 * The handler is the same for all accounts — per-request config is resolved
 * from the activity's recipient.id via the adapter-store.
 */
function registerMessageHandler(
  agentApp: InstanceType<any>,
  ActivityTypes: { Message: string },
  TurnContext: any,
  opts: {
    cfg: OpenClawConfig;
    a365Cfg: A365Config;
    runtime: RuntimeEnv;
    accountId: string;
  },
): void {
  const core = getA365Runtime();
  const log = core.logging.getChildLogger({ name: `a365:${opts.accountId}` });
  const { cfg, a365Cfg, runtime, accountId } = opts;

  // Resolve per-account config
  const acctCfg = resolveAccountA365Config(a365Cfg, accountId) ?? a365Cfg;

  // Handle welcome message
  agentApp.onConversationUpdate("membersAdded", async (context: typeof TurnContext) => {
    log.debug("members added event");
    const welcomeMessage = acctCfg?.welcomeMessage;
    if (welcomeMessage !== undefined && welcomeMessage !== "") {
      await context.sendActivity(welcomeMessage);
    }
  });

  // Handle all messages
  agentApp.onActivity(
    ActivityTypes.Message,
    async (context: typeof TurnContext, _state: unknown) => {
      const activity = context.activity;
      const text = activity.text?.trim();

      if (!text) {
        log.debug("skipping empty message");
        return;
      }

      const metadata = extractMessageMetadata(activity);
      log.info("received message", {
        from: metadata.userName || metadata.userId,
        isGroup: metadata.isGroup,
        textLength: text.length,
      });

      // Store conversation reference with accountId for proactive messaging.
      try {
        const convRef = activity.getConversationReference();
        log.info("Saving conversation reference", {
          conversationId: convRef.conversation?.id,
          serviceUrl: convRef.serviceUrl,
          agentRole: convRef.agent?.role,
          accountId,
        });
        await saveConversationReference(convRef, metadata.userAadId, accountId);
        log.info("Conversation reference saved successfully");
      } catch (err) {
        log.error(`Failed to save conversation reference: ${String(err)}`);
      }

      // Check allowlist if configured
      const allowFrom = acctCfg?.allowFrom;
      if (allowFrom && allowFrom.length > 0 && !allowFrom.includes("*")) {
        const userAllowed =
          allowFrom.includes(metadata.userId) ||
          allowFrom.includes(metadata.userEmail || "") ||
          allowFrom.includes(metadata.userAadId || "");
        if (!userAllowed) {
          log.debug("user not in allowlist", { userId: metadata.userId });
          return;
        }
      }

      // Determine user role based on owner config
      const isOwner =
        (acctCfg?.owner &&
          metadata.userEmail?.toLowerCase() === acctCfg.owner.toLowerCase()) ||
        (acctCfg?.ownerAadId && metadata.userAadId === acctCfg.ownerAadId);
      const userRole = isOwner ? "Owner" : "Requester";

      // Run the message handler within the Graph tool context for thread-safety
      await runWithGraphToolContext(
        {
          agentIdentity: acctCfg?.agentIdentity || acctCfg?.owner,
          currentUserEmail: metadata.userEmail,
          currentUserAadId: metadata.userAadId,
          currentUserRole: userRole,
          sendActivity: async (activity) => {
            const result = await context.sendActivity(activity);
            return { id: result?.id };
          },
        },
        async () => {
          const senderId = metadata.userAadId || metadata.userId;
          const conversationId = metadata.conversationId;
          const isDirectMessage = !metadata.isGroup;

          const route = core.channel.routing.resolveAgentRoute({
            cfg,
            channel: "a365",
            peer: {
              kind: isDirectMessage ? "dm" : "group",
              id: isDirectMessage ? senderId : conversationId,
            },
          });

          const a365From = isDirectMessage
            ? `a365:${senderId}`
            : `a365:group:${conversationId}`;
          const a365To = isDirectMessage ? `user:${senderId}` : `conversation:${conversationId}`;

          const ctxPayload = core.channel.reply.finalizeInboundContext({
            Body: text,
            RawBody: text,
            CommandBody: text,
            From: a365From,
            To: a365To,
            SessionKey: route.sessionKey,
            AccountId: route.accountId,
            ChatType: isDirectMessage ? "direct" : "group",
            ConversationLabel: metadata.userName || senderId,
            SenderName: metadata.userName || senderId,
            SenderId: senderId,
            Provider: "a365" as const,
            Surface: "a365" as const,
            MessageSid: metadata.activityId,
            Timestamp: Date.now(),
            WasMentioned: true,
            CommandAuthorized: isOwner,
            OriginatingChannel: "a365" as const,
            OriginatingTo: conversationId,
          });

          let replyCount = 0;
          const pendingSends: Promise<void>[] = [];

          const sendReply = async (replyText: string) => {
            try {
              log.debug("sendReply called", { length: replyText.length });
              const result = await context.sendActivity(replyText);
              replyCount++;
              log.debug("reply sent successfully", { replyCount, resultId: result?.id });
            } catch (sendErr) {
              const err = sendErr as Error;
              log.error("sendActivity failed", { error: err?.message });
            }
          };

          const queuedCounts = { tool: 0, block: 0, final: 0 };
          const dispatcher = {
            sendToolResult: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.tool++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            sendBlockReply: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.block++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            sendFinalReply: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.final++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            waitForIdle: async () => {
              await Promise.all(pendingSends);
            },
            getQueuedCounts: () => queuedCounts,
          };

          const replyOptions = {
            onReplyStart: async () => {
              try {
                log.debug("sending typing indicator");
                await context.sendActivity({ type: "typing" });
                log.debug("typing indicator sent");
              } catch (typingErr) {
                const err = typingErr as Error;
                log.debug("typing indicator failed", {
                  error: String(err),
                  message: err?.message,
                  stack: err?.stack,
                });
              }
            },
            onTypingController: () => {},
            onTypingCleanup: () => {},
          };

          try {
            log.info("dispatching to agent", { sessionKey: route.sessionKey });

            const { queuedFinal, counts } = await core.channel.reply.dispatchReplyFromConfig({
              ctx: ctxPayload,
              cfg,
              dispatcher,
              replyOptions,
            });

            await Promise.all(pendingSends);

            log.info("dispatch complete", { queuedFinal, textCount: counts?.text ?? 0, repliesSent: replyCount });

            // Update the main session's lastChannel/lastTo for cron delivery support.
            try {
              const storePath = core.channel.session.resolveStorePath(cfg.session?.store);
              const mainSessionKey = "agent:main:main";
              await core.channel.session.updateLastRoute({
                storePath,
                sessionKey: mainSessionKey,
                channel: "a365",
                to: conversationId,
              });
              log.info("Updated main session for cron delivery", { conversationId });
            } catch (updateErr) {
              log.error(`Failed to update main session: ${String(updateErr)}`);
            }
          } catch (err) {
            log.error("handler failed", { error: String(err) });
            runtime.error?.(`a365 handler failed: ${String(err)}`);

            try {
              await context.sendActivity(
                "I encountered an error processing your message. Please try again.",
              );
            } catch {
              // Ignore send failure
            }
          }
        },
      );
    },
  );
}

/**
 * Start the A365 Microsoft Agents provider.
 *
 * Multi-account mode: creates one AgentApplication + CloudAdapter per account,
 * routes inbound activities by recipient.id via a shared Express server.
 *
 * Single-account mode (backward compat): falls back to a single adapter.
 */
export async function monitorA365Provider(opts: MonitorA365Opts): Promise<MonitorA365Result> {
  const core = getA365Runtime();
  const log = core.logging.getChildLogger({ name: "a365" });
  const cfg = opts.cfg;
  const a365Cfg = cfg.channels?.a365 as A365Config | undefined;

  if (!a365Cfg?.enabled) {
    log.debug("a365 provider disabled");
    return { app: null, shutdown: async () => {} };
  }

  const runtime: RuntimeEnv = opts.runtime ?? {
    log: console.log,
    error: console.error,
    exit: (code: number): never => {
      throw new Error(`exit ${code}`);
    },
  };

  const port = a365Cfg.webhook?.port ?? 3978;

  // Guard: prevent double-start
  if (a365ServerActive) {
    log.warn(`a365 server already active on port ${port}, skipping duplicate start`);
    await new Promise<void>((resolve) => {
      if (opts.abortSignal) {
        opts.abortSignal.addEventListener("abort", () => resolve());
      }
    });
    return { app: null, shutdown: async () => {} };
  }

  // Determine accounts to initialize
  const accounts = a365Cfg.accounts;
  const hasMultiAccounts = accounts && Object.keys(accounts).length > 0;

  // Dynamic imports for Microsoft Agents SDK
  const { AgentApplication, MemoryStorage, TurnContext, TurnState, CloudAdapter } = await import(
    "@microsoft/agents-hosting"
  );
  const { ActivityTypes } = await import("@microsoft/agents-activity");

  type ApplicationTurnState = typeof TurnState;

  if (hasMultiAccounts) {
    // --- Multi-account mode ---
    log.info(`starting a365 provider in multi-account mode (${Object.keys(accounts).length} accounts, port ${port})`);

    const agentApps: Array<{ accountId: string; agentApp: InstanceType<typeof AgentApplication>; appId: string }> = [];

    // Create adapters SEQUENTIALLY (env vars are process-global, SDK caches at creation time)
    for (const [accountId, acctConfig] of Object.entries(accounts)) {
      if (acctConfig.enabled === false) {
        log.info(`skipping disabled account: ${accountId}`);
        continue;
      }

      const acctCfg = resolveAccountA365Config(a365Cfg, accountId);
      const creds = resolveA365Credentials(acctCfg);
      if (!creds) {
        log.warn(`skipping account ${accountId}: no credentials configured`);
        continue;
      }

      log.info(`creating adapter for account: ${accountId} (appId: ${creds.appId})`);

      // Set env vars for this account BEFORE creating the adapter
      setEnvForAccount(creds, port);

      // Create AgentApplication
      const storage = new MemoryStorage();
      const agentApp = new AgentApplication<ApplicationTurnState>({ storage });

      // Register message handler for this account
      registerMessageHandler(agentApp, ActivityTypes, TurnContext, {
        cfg,
        a365Cfg,
        runtime,
        accountId,
      });

      // Extract adapter from the AgentApplication
      const adapter = (agentApp.adapter ?? new CloudAdapter()) as InstanceType<typeof CloudAdapter>;

      // Resolve blueprint client ID for this account
      const blueprintClientId =
        acctCfg?.graph?.blueprintClientAppId?.trim() ||
        process.env.BLUEPRINT_CLIENT_APP_ID?.trim() ||
        creds.appId;

      // Store in adapter-store
      registerAdapter({
        accountId,
        appId: creds.appId,
        adapter,
        blueprintClientId,
        agentApp,
      });

      agentApps.push({ accountId, agentApp, appId: creds.appId });
      log.info(`adapter registered for account: ${accountId} (blueprintClientId: ${blueprintClientId})`);
    }

    if (agentApps.length === 0) {
      log.error("no accounts configured with valid credentials");
      return { app: null, shutdown: async () => {} };
    }

    // Start custom Express server for multi-adapter routing
    const { default: express } = await import("express");
    const app = express();
    app.use(express.json());

    // POST /api/messages — route to correct adapter by activity.recipient.id
    // adapter.process() expects (req, res, logic) where logic is (context) => Promise<void>.
    // AgentApplication.run(context) is the standard entry point.
    app.post("/api/messages", (req: any, res: any) => {
      const activity = req.body;
      const recipientId = activity?.recipient?.id;

      if (!recipientId) {
        log.warn("inbound activity missing recipient.id, using first adapter");
      }

      const entry = recipientId ? getAdapterByRecipientId(recipientId) : null;
      if (entry) {
        const accountId = resolveAccountIdByRecipientId(recipientId);
        log.debug("routing activity", { recipientId, accountId });
        const match = agentApps.find((a) => a.accountId === accountId);
        if (match) {
          entry.adapter.process(req, res, async (context: any) => match.agentApp.run(context));
          return;
        }
      }

      // Fallback: use first adapter
      const fallback = agentApps[0];
      log.debug("routing activity to default adapter", { recipientId, accountId: fallback.accountId });
      const fallbackEntry = getAdapterByRecipientId(fallback.appId);
      if (fallbackEntry) {
        fallbackEntry.adapter.process(req, res, async (context: any) => fallback.agentApp.run(context));
      } else {
        res.status(500).json({ error: "No adapter available" });
      }
    });

    // GET /api/health
    app.get("/api/health", (_req: any, res: any) => {
      res.json({
        ok: true,
        accounts: agentApps.map((a) => a.accountId),
        uptime: process.uptime(),
      });
    });

    a365ServerActive = true;

    // Start listening
    await new Promise<void>((resolve, reject) => {
      const server = app.listen(port, () => {
        log.info(`a365 multi-account server listening on port ${port} (${agentApps.length} adapters)`);
        resolve();
      });
      server.on("error", (err: Error) => {
        log.error(`failed to start server: ${err.message}`);
        reject(err);
      });

      // Clean up on abort
      if (opts.abortSignal) {
        opts.abortSignal.addEventListener("abort", () => {
          server.close();
          a365ServerActive = false;
        });
      }
    });

    const shutdown = async () => {
      log.info("shutting down a365 multi-account provider");
      a365ServerActive = false;
    };

    // Block until abort signal
    await new Promise<void>((resolve) => {
      if (opts.abortSignal) {
        opts.abortSignal.addEventListener("abort", () => {
          void shutdown();
          resolve();
        });
      }
    });

    return { app: agentApps, shutdown };
  }

  // --- Single-account mode (backward compat) ---
  log.info(`starting a365 provider in single-account mode (port ${port})`);

  const creds = resolveA365Credentials(a365Cfg);
  if (!creds) {
    log.error("A365 credentials not configured - set appId/appPassword/tenantId in config or A365_APP_ID/A365_APP_PASSWORD/A365_TENANT_ID env vars");
    return { app: null, shutdown: async () => {} };
  }

  setEnvForAccount(creds, port);

  const storage = new MemoryStorage();
  const agentApp = new AgentApplication<ApplicationTurnState>({ storage });

  // Register message handler for single-account mode
  registerMessageHandler(agentApp, ActivityTypes, TurnContext, {
    cfg,
    a365Cfg,
    runtime,
    accountId: "__default__",
  });

  // Store the adapter for proactive messaging
  const adapter = (agentApp.adapter ?? new CloudAdapter()) as InstanceType<typeof CloudAdapter>;
  setAdapter(adapter);

  const blueprintClientId =
    a365Cfg?.graph?.blueprintClientAppId?.trim() ||
    process.env.BLUEPRINT_CLIENT_APP_ID?.trim() ||
    creds.appId;
  setBlueprintClientId(blueprintClientId);
  log.info("Stored adapter and blueprint client ID for proactive messaging", { blueprintClientId });

  // Start the server using the Agents SDK
  const { startServer } = await import("@microsoft/agents-hosting-express");

  a365ServerActive = true;

  await startServer(agentApp);

  log.info(`a365 provider started on port ${port}`);

  const shutdown = async () => {
    log.info("shutting down a365 provider");
    a365ServerActive = false;
  };

  // Block until abort signal
  await new Promise<void>((resolve) => {
    if (opts.abortSignal) {
      opts.abortSignal.addEventListener("abort", () => {
        void shutdown();
        resolve();
      });
    }
  });

  return { app: agentApp, shutdown };
}
