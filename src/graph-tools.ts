import { AsyncLocalStorage } from "node:async_hooks";
import { Type, type TSchema } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import type { A365Config, GraphCalendarEvent, GraphMailMessage, GraphMailFolder } from "./types.js";
import { getGraphToken } from "./token.js";
import { getA365Runtime } from "./runtime.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_BETA_URL = "https://graph.microsoft.com/beta";
const DEFAULT_TIMEZONE = "UTC";

/**
 * Format a Graph resource ID for use in URL path segments.
 * Returns the raw ID prefixed with / — no encoding.
 * This matches the original working implementation.
 */
function graphId(id: string): string {
  return `/${id}`;
}


/**
 * Get the default timezone from config, falling back to UTC.
 */
function getDefaultTimezone(cfg?: A365Config): string {
  return cfg?.businessHours?.timezone || DEFAULT_TIMEZONE;
}

/**
 * Get the logger for Graph API operations.
 * Returns a no-op logger if runtime is not yet initialized.
 */
function getLogger() {
  try {
    return getA365Runtime().logging.getChildLogger({ name: "a365-graph" });
  } catch {
    return {
      debug: () => {},
      info: () => {},
      warn: () => {},
      error: () => {},
    };
  }
}

/**
 * Context for Graph API tool execution.
 * This provides user information for token acquisition.
 */
export type GraphToolContext = {
  /** Username (email) of the agent service account */
  agentIdentity?: string;
  /** Username (email) of the current user interacting with the agent */
  currentUserEmail?: string;
  /** AAD Object ID of the current user */
  currentUserAadId?: string;
  /** Role of the current user */
  currentUserRole?: "Owner" | "Requester";
  /** Callback to send an activity directly to the conversation (e.g. GIF attachments) */
  sendActivity?: (activity: unknown) => Promise<{ id?: string }>;
};

/**
 * AsyncLocalStorage for thread-safe tool context.
 * This ensures each request has its own isolated context,
 * preventing cross-request data leakage in concurrent scenarios.
 */
const toolContextStorage = new AsyncLocalStorage<GraphToolContext>();

/**
 * Run a function with the given tool context.
 * Use this to wrap request handlers to ensure proper context isolation.
 */
export function runWithGraphToolContext<T>(ctx: GraphToolContext, fn: () => T): T {
  return toolContextStorage.run(ctx, fn);
}

/**
 * Get the current tool context from AsyncLocalStorage.
 */
export function getGraphToolContext(): GraphToolContext | undefined {
  return toolContextStorage.getStore();
}

/**
 * @deprecated Use runWithGraphToolContext instead for thread-safe context management.
 * This function is kept for backwards compatibility but will set context globally
 * which is not safe in concurrent scenarios.
 */
export function setGraphToolContext(_ctx: GraphToolContext | undefined): void {
  // No-op: Context should be set via runWithGraphToolContext
  // This is kept for API compatibility but logs a warning
  getLogger().warn("setGraphToolContext is deprecated - use runWithGraphToolContext for thread-safe context");
}

/**
 * Make a request to the Microsoft Graph API.
 * Uses the agent username (service account) for token acquisition.
 *
 * TODO: Add retry logic with exponential backoff for transient failures (429, 503).
 */
async function graphRequest<T>(
  cfg: A365Config | undefined,
  method: string,
  path: string,
  body?: unknown,
  options?: { useBeta?: boolean },
): Promise<{ ok: true; data: T } | { ok: false; error: string; status?: number; errorCode?: string; path?: string; rawError?: string }> {
  const log = getLogger();

  // Get the username for token acquisition
  // Use agent username from context (thread-safe) or config
  const toolContext = getGraphToolContext();
  const agentIdentity =
    toolContext?.agentIdentity ||
    cfg?.agentIdentity ||
    cfg?.owner;

  if (!agentIdentity) {
    return { ok: false, error: "Agent username not configured. Set agentIdentity or owner in config." };
  }

  const token = await getGraphToken(cfg, agentIdentity);
  if (!token) {
    return { ok: false, error: "Graph API token not available. Check T1/T2/User flow configuration (blueprintClientAppId, blueprintClientSecret, aaInstanceId)." };
  }

  const baseUrl = options?.useBeta ? GRAPH_BETA_URL : GRAPH_BASE_URL;
  const url = `${baseUrl}${path}`;
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  log.debug("Graph API request", { method, path, agentIdentity });

  try {
    const response = await fetch(url, {
      method,
      headers,
      body: body ? JSON.stringify(body) : undefined,
    });

    log.debug("Graph API response", { status: response.status });

    if (!response.ok) {
      const errorText = await response.text();
      log.warn("Graph API error", { status: response.status, error: errorText.slice(0, 200) });
      let errorMessage = `Graph API error: ${response.status}`;
      let errorCode = "";
      let rawError = errorText.slice(0, 500);
      try {
        const errorJson = JSON.parse(errorText);
        errorCode = errorJson.error?.code || "";
        errorMessage = errorJson.error?.message || errorMessage;
        rawError = JSON.stringify(errorJson.error, null, 0).slice(0, 500);
      } catch {
        errorMessage = errorText || errorMessage;
      }
      return { ok: false, error: errorMessage, status: response.status, errorCode, path, rawError };
    }

    // Handle empty-body success responses (202 Accepted, 204 No Content)
    // Graph API returns 202 for async operations like sendMail, 204 for DELETE
    if (response.status === 202 || response.status === 204) {
      return { ok: true, data: {} as T };
    }

    const text = await response.text();
    if (!text) {
      return { ok: true, data: {} as T };
    }

    const data = JSON.parse(text) as T;
    return { ok: true, data };
  } catch (err) {
    log.error("Graph API network error", { error: String(err) });
    return { ok: false, error: `Network error: ${String(err)}` };
  }
}

/**
 * Well-known folder names that Graph API accepts as direct path segments.
 * These short strings won't be corrupted by an LLM.
 */
const WELL_KNOWN_FOLDERS = new Set([
  "inbox", "drafts", "sentitems", "deleteditems", "junkemail",
  "archive", "outbox", "clutter", "conflicts", "localfailures",
  "serverfailures", "syncissues",
]);

/**
 * Resolve a mail folder display name to its Graph API folder ID.
 *
 * This is the core fix for the LLM-ID-corruption problem: instead of letting
 * the LLM pass long Base64 folder IDs between tool calls (which it tends to
 * truncate or mangle), the tools accept human-readable display names and
 * resolve them to IDs internally via the Graph API.
 *
 * Resolution order:
 * 1. Well-known names (inbox, drafts, etc.) → used directly
 * 2. Top-level folder displayName match (case-insensitive)
 * 3. One-level-deep child folder search
 */
async function resolveMailFolderByName(
  cfg: A365Config | undefined,
  userId: string,
  folderName: string,
): Promise<{ ok: true; folderId: string; displayName: string } | { ok: false; error: string }> {
  const log = getLogger();

  // Well-known folder names work directly as path segments
  const lower = folderName.toLowerCase().trim();
  if (WELL_KNOWN_FOLDERS.has(lower)) {
    return { ok: true, folderId: lower, displayName: folderName };
  }

  // List top-level folders
  const path = `/users/${encodeURIComponent(userId)}/mailFolders?$top=100&$select=id,displayName,childFolderCount&includeHiddenFolders=true`;
  const result = await graphRequest<{ value: GraphMailFolder[] }>(cfg, "GET", path);
  if (!result.ok) {
    return { ok: false, error: `Failed to list folders: ${result.error}` };
  }

  // Case-insensitive match on displayName at top level
  const match = result.data.value.find(
    (f) => f.displayName?.toLowerCase() === lower,
  );
  if (match?.id) {
    log.debug("Resolved folder by name", { folderName, folderId: match.id.substring(0, 20) });
    return { ok: true, folderId: match.id, displayName: match.displayName ?? folderName };
  }

  // Search child folders (one level deep) for folders that have children
  for (const folder of result.data.value) {
    if ((folder.childFolderCount ?? 0) > 0 && folder.id) {
      const childPath = `/users/${encodeURIComponent(userId)}/mailFolders${graphId(folder.id)}/childFolders?$top=100&$select=id,displayName&includeHiddenFolders=true`;
      const childResult = await graphRequest<{ value: GraphMailFolder[] }>(cfg, "GET", childPath);
      if (childResult.ok) {
        const childMatch = childResult.data.value.find(
          (f) => f.displayName?.toLowerCase() === lower,
        );
        if (childMatch?.id) {
          log.debug("Resolved child folder by name", { folderName, parent: folder.displayName, folderId: childMatch.id.substring(0, 20) });
          return { ok: true, folderId: childMatch.id, displayName: childMatch.displayName ?? folderName };
        }
      }
    }
  }

  // Not found — return available names so the LLM can retry
  const available = result.data.value
    .map((f) => f.displayName)
    .filter(Boolean)
    .join(", ");
  return { ok: false, error: `Folder "${folderName}" not found. Available top-level folders: ${available}` };
}

/**
 * Validate common tool parameters.
 */
function validateUserId(userId: string): { ok: true } | { ok: false; error: string } {
  if (!userId || !userId.trim()) {
    return { ok: false, error: "userId is required and cannot be empty" };
  }
  // Basic email format check (loose validation - Graph API will reject invalid IDs)
  if (!userId.includes("@") && !/^[0-9a-f-]{36}$/i.test(userId)) {
    return { ok: false, error: "userId should be an email address or a valid GUID" };
  }
  return { ok: true };
}

/**
 * Validate ISO datetime string format.
 */
function validateIsoDateTime(dateTime: string, fieldName: string): { ok: true } | { ok: false; error: string } {
  if (!dateTime || !dateTime.trim()) {
    return { ok: false, error: `${fieldName} is required` };
  }
  // Basic ISO format check (YYYY-MM-DDTHH:MM:SS)
  if (!/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?/.test(dateTime)) {
    return { ok: false, error: `${fieldName} should be in ISO format (e.g., 2024-01-15T14:00:00)` };
  }
  return { ok: true };
}

/**
 * Validate email addresses in an array.
 */
function validateEmails(emails: string[], fieldName: string): { ok: true } | { ok: false; error: string } {
  for (const email of emails) {
    if (!email.includes("@")) {
      return { ok: false, error: `Invalid email address in ${fieldName}: ${email}` };
    }
  }
  return { ok: true };
}

type ToolResult = AgentToolResult<unknown>;

// --- GIF dedup ring buffer (module-level, in-memory) ---
const RECENT_GIF_MAX = 20;
const recentGifIds: number[] = [];
const recentGifSet = new Set<number>();

function addToRecentGifs(id: number): void {
  if (recentGifSet.has(id)) return;
  if (recentGifIds.length >= RECENT_GIF_MAX) {
    const evicted = recentGifIds.shift()!;
    recentGifSet.delete(evicted);
  }
  recentGifIds.push(id);
  recentGifSet.add(id);
}

function isRecentGif(id: number): boolean {
  return recentGifSet.has(id);
}

/**
 * Search Klipy for GIFs, pick a non-recent random result, and send it
 * directly into the conversation via the turn context's sendActivity.
 */
async function sendGif(
  cfg: A365Config | undefined,
  params: { query: string },
): Promise<ToolResult> {
  const log = getLogger();
  const klipyKey = cfg?.klipyApiKey || process.env.KLIPY_API_KEY;
  if (!klipyKey) {
    return { isError: true, content: [{ type: "text", text: "Klipy API key not configured. Set KLIPY_API_KEY env var or klipyApiKey in config." }] };
  }

  const ctx = getGraphToolContext();
  if (!ctx?.sendActivity) {
    return { isError: true, content: [{ type: "text", text: "Cannot send GIF: sendActivity not available in current context." }] };
  }

  const { query } = params;
  const url = `https://api.klipy.com/api/v1/${encodeURIComponent(klipyKey)}/gifs/search?q=${encodeURIComponent(query)}&per_page=20`;

  try {
    const resp = await fetch(url);
    if (!resp.ok) {
      log.warn("Klipy API error", { status: resp.status });
      return { isError: true, content: [{ type: "text", text: `Klipy API error: ${resp.status}` }] };
    }

    const json = await resp.json() as { data?: { data?: Array<{ id: number; title?: string; slug?: string; file?: { hd?: { gif?: { url?: string } } } }> } };
    const results = json.data?.data;
    if (!results || results.length === 0) {
      return { content: [{ type: "text", text: JSON.stringify({ sent: false, reason: "No GIFs found for that query." }) }] };
    }

    // Filter out recently-used GIFs
    const fresh = results.filter((g) => !isRecentGif(g.id));
    const pool = fresh.length > 0 ? fresh : results; // fall back to all if everything is recent

    // Pick a random result
    const pick = pool[Math.floor(Math.random() * pool.length)];
    const gifUrl = pick.file?.hd?.gif?.url;
    if (!gifUrl) {
      return { isError: true, content: [{ type: "text", text: "Selected GIF has no HD URL available." }] };
    }

    // Send the GIF as an inline attachment via the turn context
    await ctx.sendActivity({
      type: "message",
      attachments: [
        {
          contentType: "image/gif",
          contentUrl: gifUrl,
          name: `${pick.slug || "gif"}.gif`,
        },
      ],
    });

    addToRecentGifs(pick.id);
    log.debug("GIF sent", { id: pick.id, title: pick.title, query });

    return {
      content: [{ type: "text", text: JSON.stringify({ sent: true, title: pick.title || pick.slug || "GIF", query }) }],
    };
  } catch (err) {
    log.error("sendGif failed", { error: String(err) });
    return { isError: true, content: [{ type: "text", text: `Failed to send GIF: ${String(err)}` }] };
  }
}

/**
 * Get calendar events for a user within a date range.
 */
async function getCalendarEvents(
  cfg: A365Config | undefined,
  params: { userId: string; startDate: string; endDate: string },
): Promise<ToolResult> {
  const { userId, startDate, endDate } = params;

  // Validate inputs
  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  const startCheck = validateIsoDateTime(startDate, "startDate");
  if (!startCheck.ok) return { isError: true, content: [{ type: "text", text: startCheck.error }] };

  const endCheck = validateIsoDateTime(endDate, "endDate");
  if (!endCheck.ok) return { isError: true, content: [{ type: "text", text: endCheck.error }] };

  const path = `/users/${encodeURIComponent(userId)}/calendar/calendarView?startDateTime=${encodeURIComponent(startDate)}&endDateTime=${encodeURIComponent(endDate)}&$orderby=start/dateTime&$top=50`;

  const result = await graphRequest<{ value: GraphCalendarEvent[] }>(cfg, "GET", path);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  const events = result.data.value.map((event) => ({
    id: event.id,
    subject: event.subject,
    start: event.start,
    end: event.end,
    location: event.location?.displayName,
    attendees: event.attendees?.map((a) => ({
      email: a.emailAddress.address,
      name: a.emailAddress.name,
      response: a.status?.response,
    })),
    isOnlineMeeting: event.isOnlineMeeting,
    onlineMeetingUrl: event.onlineMeetingUrl,
    showAs: event.showAs,
    isCancelled: event.isCancelled,
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ events, count: events.length }, null, 2),
      },
    ],
  };
}

/**
 * Create a calendar event for a user.
 */
async function createCalendarEvent(
  cfg: A365Config | undefined,
  params: {
    userId: string;
    subject: string;
    startDateTime: string;
    endDateTime: string;
    timeZone?: string;
    attendees?: string[];
    location?: string;
    body?: string;
    isOnlineMeeting?: boolean;
  },
): Promise<ToolResult> {
  const defaultTz = getDefaultTimezone(cfg);
  const {
    userId,
    subject,
    startDateTime,
    endDateTime,
    timeZone = defaultTz,
    attendees = [],
    location,
    body,
    isOnlineMeeting = false,
  } = params;

  // Validate inputs
  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  const startCheck = validateIsoDateTime(startDateTime, "startDateTime");
  if (!startCheck.ok) return { isError: true, content: [{ type: "text", text: startCheck.error }] };

  const endCheck = validateIsoDateTime(endDateTime, "endDateTime");
  if (!endCheck.ok) return { isError: true, content: [{ type: "text", text: endCheck.error }] };

  if (attendees.length > 0) {
    const emailsCheck = validateEmails(attendees, "attendees");
    if (!emailsCheck.ok) return { isError: true, content: [{ type: "text", text: emailsCheck.error }] };
  }

  const path = `/users/${encodeURIComponent(userId)}/calendar/events`;

  const eventBody: Partial<GraphCalendarEvent> = {
    subject,
    start: { dateTime: startDateTime, timeZone },
    end: { dateTime: endDateTime, timeZone },
    isOnlineMeeting,
  };

  if (attendees.length > 0) {
    eventBody.attendees = attendees.map((email) => ({
      emailAddress: { address: email },
      type: "required",
    }));
  }

  if (location) {
    eventBody.location = { displayName: location };
  }

  if (body) {
    eventBody.body = { contentType: "text", content: body };
  }

  const result = await graphRequest<GraphCalendarEvent>(cfg, "POST", path, eventBody);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            success: true,
            eventId: result.data.id,
            subject: result.data.subject,
            start: result.data.start,
            end: result.data.end,
            onlineMeetingUrl: result.data.onlineMeetingUrl,
          },
          null,
          2,
        ),
      },
    ],
  };
}

/**
 * Update an existing calendar event.
 */
async function updateCalendarEvent(
  cfg: A365Config | undefined,
  params: {
    userId: string;
    eventId: string;
    subject?: string;
    startDateTime?: string;
    endDateTime?: string;
    timeZone?: string;
    attendees?: string[];
    location?: string;
    body?: string;
    isOnlineMeeting?: boolean;
  },
): Promise<ToolResult> {
  const defaultTz = getDefaultTimezone(cfg);
  const { userId, eventId, subject, startDateTime, endDateTime, timeZone, attendees, location, body, isOnlineMeeting } = params;

  if (attendees && attendees.length > 0) {
    const emailsCheck = validateEmails(attendees, "attendees");
    if (!emailsCheck.ok) return { isError: true, content: [{ type: "text", text: emailsCheck.error }] };
  }

  const path = `/users/${encodeURIComponent(userId)}/calendar/events${graphId(eventId)}`;

  const eventBody: Partial<GraphCalendarEvent> = {};

  if (subject !== undefined) {
    eventBody.subject = subject;
  }
  if (startDateTime !== undefined) {
    eventBody.start = { dateTime: startDateTime, timeZone: timeZone || defaultTz };
  }
  if (endDateTime !== undefined) {
    eventBody.end = { dateTime: endDateTime, timeZone: timeZone || defaultTz };
  }
  if (attendees !== undefined) {
    eventBody.attendees = attendees.map((email) => ({
      emailAddress: { address: email },
      type: "required",
    }));
  }
  if (location !== undefined) {
    eventBody.location = { displayName: location };
  }
  if (body !== undefined) {
    eventBody.body = { contentType: "text", content: body };
  }
  if (isOnlineMeeting !== undefined) {
    eventBody.isOnlineMeeting = isOnlineMeeting;
  }

  const result = await graphRequest<GraphCalendarEvent>(cfg, "PATCH", path, eventBody);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            success: true,
            eventId: result.data.id,
            subject: result.data.subject,
            start: result.data.start,
            end: result.data.end,
          },
          null,
          2,
        ),
      },
    ],
  };
}

/**
 * Delete a calendar event.
 */
async function deleteCalendarEvent(
  cfg: A365Config | undefined,
  params: { userId: string; eventId: string },
): Promise<ToolResult> {
  const { userId, eventId } = params;

  const path = `/users/${encodeURIComponent(userId)}/calendar/events${graphId(eventId)}`;

  const result = await graphRequest<unknown>(cfg, "DELETE", path);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ success: true, deleted: eventId }, null, 2),
      },
    ],
  };
}

/**
 * Send an email using Microsoft Graph.
 */
async function sendEmail(
  cfg: A365Config | undefined,
  params: {
    userId: string;
    to: string[];
    subject: string;
    body: string;
    cc?: string[];
    bcc?: string[];
    importance?: "low" | "normal" | "high";
  },
): Promise<ToolResult> {
  const { userId, to, subject, body, cc = [], bcc = [], importance = "normal" } = params;

  // Validate inputs
  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (to.length === 0) {
    return { isError: true, content: [{ type: "text", text: "At least one recipient is required in 'to' field" }] };
  }

  const toCheck = validateEmails(to, "to");
  if (!toCheck.ok) return { isError: true, content: [{ type: "text", text: toCheck.error }] };

  if (cc.length > 0) {
    const ccCheck = validateEmails(cc, "cc");
    if (!ccCheck.ok) return { isError: true, content: [{ type: "text", text: ccCheck.error }] };
  }

  if (bcc.length > 0) {
    const bccCheck = validateEmails(bcc, "bcc");
    if (!bccCheck.ok) return { isError: true, content: [{ type: "text", text: bccCheck.error }] };
  }

  const path = `/users/${encodeURIComponent(userId)}/sendMail`;

  const mailBody = {
    message: {
      subject,
      body: {
        contentType: "text",
        content: body,
      },
      toRecipients: to.map((email) => ({ emailAddress: { address: email } })),
      ccRecipients: cc.map((email) => ({ emailAddress: { address: email } })),
      bccRecipients: bcc.map((email) => ({ emailAddress: { address: email } })),
      importance,
    },
    saveToSentItems: true,
  };

  const result = await graphRequest<unknown>(cfg, "POST", path, mailBody);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            success: true,
            message: `Email sent successfully to ${to.join(", ")}`,
          },
          null,
          2,
        ),
      },
    ],
  };
}

/**
 * Get user information from Microsoft Graph.
 */
async function getUserInfo(
  cfg: A365Config | undefined,
  params: { userId: string },
): Promise<ToolResult> {
  const { userId } = params;

  const path = `/users/${encodeURIComponent(userId)}?$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation`;

  const result = await graphRequest<{
    id: string;
    displayName: string;
    mail: string;
    userPrincipalName: string;
    jobTitle?: string;
    department?: string;
    officeLocation?: string;
  }>(cfg, "GET", path);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(result.data, null, 2),
      },
    ],
  };
}

/**
 * Find available meeting times using findMeetingTimes API.
 */
async function findMeetingTimes(
  cfg: A365Config | undefined,
  params: {
    userId: string;
    attendees: string[];
    durationMinutes: number;
    startDateTime: string;
    endDateTime: string;
    timeZone?: string;
  },
): Promise<ToolResult> {
  const defaultTz = getDefaultTimezone(cfg);
  const {
    userId,
    attendees,
    durationMinutes,
    startDateTime,
    endDateTime,
    timeZone = defaultTz,
  } = params;

  const path = `/users/${encodeURIComponent(userId)}/findMeetingTimes`;

  const body = {
    attendees: attendees.map((email) => ({
      emailAddress: { address: email },
      type: "required",
    })),
    timeConstraint: {
      activityDomain: "work",
      timeSlots: [
        {
          start: { dateTime: startDateTime, timeZone },
          end: { dateTime: endDateTime, timeZone },
        },
      ],
    },
    meetingDuration: `PT${durationMinutes}M`,
    maxCandidates: 5,
    isOrganizerOptional: false,
    returnSuggestionReasons: true,
  };

  const result = await graphRequest<{
    meetingTimeSuggestions: Array<{
      meetingTimeSlot: {
        start: { dateTime: string; timeZone: string };
        end: { dateTime: string; timeZone: string };
      };
      confidence: number;
      organizerAvailability: string;
      attendeeAvailability: Array<{
        attendee: { emailAddress: { address: string } };
        availability: string;
      }>;
      suggestionReason?: string;
    }>;
    emptySuggestionsReason?: string;
  }>(cfg, "POST", path, body);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  const suggestions = result.data.meetingTimeSuggestions.map((s) => ({
    start: s.meetingTimeSlot.start,
    end: s.meetingTimeSlot.end,
    confidence: s.confidence,
    organizerAvailability: s.organizerAvailability,
    attendeeAvailability: s.attendeeAvailability?.map((a) => ({
      email: a.attendee.emailAddress.address,
      availability: a.availability,
    })),
    reason: s.suggestionReason,
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            suggestions,
            count: suggestions.length,
            emptySuggestionsReason: result.data.emptySuggestionsReason,
          },
          null,
          2,
        ),
      },
    ],
  };
}

/**
 * Get emails from a user's mailbox or specific folder.
 */
async function getEmails(
  cfg: A365Config | undefined,
  params: { userId: string; folderName?: string; top?: number; filter?: string; orderBy?: string },
): Promise<ToolResult> {
  const { userId, folderName = "inbox", top = 10, filter, orderBy } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  // Resolve folder name → ID internally
  const resolved = await resolveMailFolderByName(cfg, userId, folderName);
  if (!resolved.ok) return { isError: true, content: [{ type: "text", text: resolved.error }] };

  const clampedTop = Math.min(Math.max(top, 1), 50);
  let path = `/users/${encodeURIComponent(userId)}/mailFolders${graphId(resolved.folderId)}/messages?$top=${clampedTop}&$select=id,subject,bodyPreview,from,receivedDateTime,isRead,hasAttachments,importance,flag`;

  if (orderBy) {
    path += `&$orderby=${encodeURIComponent(orderBy)}`;
  } else {
    path += `&$orderby=receivedDateTime desc`;
  }

  if (filter) {
    path += `&$filter=${encodeURIComponent(filter)}`;
  }

  const result = await graphRequest<{ value: GraphMailMessage[] }>(cfg, "GET", path);

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  const messages = result.data.value.map((msg) => ({
    id: msg.id,
    subject: msg.subject,
    preview: msg.bodyPreview,
    from: msg.from?.emailAddress ? `${msg.from.emailAddress.name || ""} <${msg.from.emailAddress.address}>` : undefined,
    receivedDateTime: msg.receivedDateTime,
    isRead: msg.isRead,
    hasAttachments: msg.hasAttachments,
    importance: msg.importance,
    flagStatus: msg.flag?.flagStatus,
  }));

  return {
    content: [{ type: "text", text: JSON.stringify({ messages, count: messages.length }, null, 2) }],
  };
}

/**
 * Read a single email message with full content.
 */
async function readEmail(
  cfg: A365Config | undefined,
  params: { userId: string; messageId: string },
): Promise<ToolResult> {
  const { userId, messageId } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (!messageId?.trim()) {
    return { isError: true, content: [{ type: "text", text: "messageId is required" }] };
  }

  const path = `/users/${encodeURIComponent(userId)}/messages${graphId(messageId)}?$select=id,subject,body,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,conversationId`;

  const result = await graphRequest<GraphMailMessage>(cfg, "GET", path);

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  const msg = result.data;
  return {
    content: [{
      type: "text",
      text: JSON.stringify({
        id: msg.id,
        subject: msg.subject,
        from: msg.from?.emailAddress ? `${msg.from.emailAddress.name || ""} <${msg.from.emailAddress.address}>` : undefined,
        to: msg.toRecipients?.map((r) => `${r.emailAddress.name || ""} <${r.emailAddress.address}>`),
        cc: msg.ccRecipients?.map((r) => `${r.emailAddress.name || ""} <${r.emailAddress.address}>`),
        body: msg.body?.content,
        bodyType: msg.body?.contentType,
        receivedDateTime: msg.receivedDateTime,
        sentDateTime: msg.sentDateTime,
        isRead: msg.isRead,
        hasAttachments: msg.hasAttachments,
        importance: msg.importance,
        conversationId: msg.conversationId,
      }, null, 2),
    }],
  };
}

/**
 * Search emails using KQL (Keyword Query Language).
 */
async function searchEmails(
  cfg: A365Config | undefined,
  params: { userId: string; query: string; top?: number },
): Promise<ToolResult> {
  const { userId, query, top = 10 } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (!query?.trim()) {
    return { isError: true, content: [{ type: "text", text: "query is required" }] };
  }

  const clampedTop = Math.min(Math.max(top, 1), 50);
  const path = `/users/${encodeURIComponent(userId)}/messages?$search="${encodeURIComponent(query)}"&$top=${clampedTop}&$select=id,subject,bodyPreview,from,receivedDateTime,isRead,hasAttachments,importance`;

  const result = await graphRequest<{ value: GraphMailMessage[] }>(cfg, "GET", path);

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  const messages = result.data.value.map((msg) => ({
    id: msg.id,
    subject: msg.subject,
    preview: msg.bodyPreview,
    from: msg.from?.emailAddress ? `${msg.from.emailAddress.name || ""} <${msg.from.emailAddress.address}>` : undefined,
    receivedDateTime: msg.receivedDateTime,
    isRead: msg.isRead,
    hasAttachments: msg.hasAttachments,
    importance: msg.importance,
  }));

  return {
    content: [{ type: "text", text: JSON.stringify({ messages, count: messages.length, query }, null, 2) }],
  };
}

/**
 * Move an email to a different folder.
 */
async function moveEmail(
  cfg: A365Config | undefined,
  params: { userId: string; messageId: string; destinationFolderName: string },
): Promise<ToolResult> {
  const { userId, messageId, destinationFolderName } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (!messageId?.trim()) {
    return { isError: true, content: [{ type: "text", text: "messageId is required" }] };
  }
  if (!destinationFolderName?.trim()) {
    return { isError: true, content: [{ type: "text", text: "destinationFolderName is required" }] };
  }

  // Resolve destination folder name → ID internally
  const resolved = await resolveMailFolderByName(cfg, userId, destinationFolderName);
  if (!resolved.ok) return { isError: true, content: [{ type: "text", text: resolved.error }] };

  const path = `/users/${encodeURIComponent(userId)}/messages${graphId(messageId)}/move`;
  const result = await graphRequest<GraphMailMessage>(cfg, "POST", path, { destinationId: resolved.folderId });

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  return {
    content: [{ type: "text", text: JSON.stringify({ success: true, newMessageId: result.data.id }, null, 2) }],
  };
}

/**
 * Delete an email message.
 */
async function deleteEmail(
  cfg: A365Config | undefined,
  params: { userId: string; messageId: string },
): Promise<ToolResult> {
  const { userId, messageId } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (!messageId?.trim()) {
    return { isError: true, content: [{ type: "text", text: "messageId is required" }] };
  }

  const path = `/users/${encodeURIComponent(userId)}/messages${graphId(messageId)}`;
  const result = await graphRequest<unknown>(cfg, "DELETE", path);

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  return {
    content: [{ type: "text", text: JSON.stringify({ success: true, deleted: messageId }, null, 2) }],
  };
}

/**
 * Mark an email as read or unread.
 */
async function markEmailRead(
  cfg: A365Config | undefined,
  params: { userId: string; messageId: string; isRead: boolean },
): Promise<ToolResult> {
  const { userId, messageId, isRead } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (!messageId?.trim()) {
    return { isError: true, content: [{ type: "text", text: "messageId is required" }] };
  }

  const path = `/users/${encodeURIComponent(userId)}/messages${graphId(messageId)}`;
  const result = await graphRequest<GraphMailMessage>(cfg, "PATCH", path, { isRead });

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  return {
    content: [{ type: "text", text: JSON.stringify({ success: true, messageId, isRead }, null, 2) }],
  };
}

/**
 * Get mail folders for a user.
 */
async function getMailFolders(
  cfg: A365Config | undefined,
  params: { userId: string; parentFolderName?: string },
): Promise<ToolResult> {
  const { userId, parentFolderName } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  let basePath: string;
  if (parentFolderName) {
    const resolved = await resolveMailFolderByName(cfg, userId, parentFolderName);
    if (!resolved.ok) return { isError: true, content: [{ type: "text", text: resolved.error }] };
    basePath = `/users/${encodeURIComponent(userId)}/mailFolders${graphId(resolved.folderId)}/childFolders`;
  } else {
    basePath = `/users/${encodeURIComponent(userId)}/mailFolders`;
  }

  const path = `${basePath}?$top=100&$select=id,displayName,parentFolderId,unreadItemCount,totalItemCount,childFolderCount&includeHiddenFolders=true`;
  const result = await graphRequest<{ value: GraphMailFolder[] }>(cfg, "GET", path);

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  const folders = result.data.value.map((f) => ({
    displayName: f.displayName,
    unreadItemCount: f.unreadItemCount,
    totalItemCount: f.totalItemCount,
    childFolderCount: f.childFolderCount,
  }));

  return {
    content: [{ type: "text", text: JSON.stringify({ folders, count: folders.length, parentFolder: parentFolderName ?? "root" }, null, 2) }],
  };
}

async function createMailFolder(
  cfg: A365Config | undefined,
  params: { userId: string; displayName: string; parentFolderName?: string },
): Promise<ToolResult> {
  const { userId, displayName, parentFolderName } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  let basePath: string;
  if (parentFolderName) {
    const resolved = await resolveMailFolderByName(cfg, userId, parentFolderName);
    if (!resolved.ok) return { isError: true, content: [{ type: "text", text: resolved.error }] };
    basePath = `/users/${encodeURIComponent(userId)}/mailFolders${graphId(resolved.folderId)}/childFolders`;
  } else {
    basePath = `/users/${encodeURIComponent(userId)}/mailFolders`;
  }

  const result = await graphRequest<GraphMailFolder>(cfg, "POST", basePath, { displayName });

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  return {
    content: [{ type: "text", text: JSON.stringify({ created: true, displayName: result.data.displayName }, null, 2) }],
  };
}

async function renameMailFolder(
  cfg: A365Config | undefined,
  params: { userId: string; folderName: string; newName: string },
): Promise<ToolResult> {
  const { userId, folderName, newName } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  // Resolve folder name → ID internally (bypasses LLM ID corruption)
  const resolved = await resolveMailFolderByName(cfg, userId, folderName);
  if (!resolved.ok) return { isError: true, content: [{ type: "text", text: resolved.error }] };

  const path = `/users/${encodeURIComponent(userId)}/mailFolders${graphId(resolved.folderId)}`;
  const result = await graphRequest<GraphMailFolder>(cfg, "PATCH", path, { displayName: newName });

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  return {
    content: [{ type: "text", text: JSON.stringify({ renamed: true, oldName: folderName, newName: result.data.displayName }, null, 2) }],
  };
}

async function deleteMailFolder(
  cfg: A365Config | undefined,
  params: { userId: string; folderName: string },
): Promise<ToolResult> {
  const { userId, folderName } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (!folderName?.trim()) {
    return { isError: true, content: [{ type: "text", text: "folderName is required" }] };
  }

  // Resolve folder name → ID internally (bypasses LLM ID corruption)
  const resolved = await resolveMailFolderByName(cfg, userId, folderName);
  if (!resolved.ok) return { isError: true, content: [{ type: "text", text: resolved.error }] };

  const path = `/users/${encodeURIComponent(userId)}/mailFolders${graphId(resolved.folderId)}`;
  const result = await graphRequest<Record<string, never>>(cfg, "DELETE", path);

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  return {
    content: [{ type: "text", text: JSON.stringify({ deleted: true, folderName }, null, 2) }],
  };
}

async function moveMailFolder(
  cfg: A365Config | undefined,
  params: { userId: string; folderName: string; destinationName: string },
): Promise<ToolResult> {
  const { userId, folderName, destinationName } = params;

  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (!folderName?.trim()) {
    return { isError: true, content: [{ type: "text", text: "folderName is required" }] };
  }
  if (!destinationName?.trim()) {
    return { isError: true, content: [{ type: "text", text: "destinationName is required" }] };
  }

  // Resolve both folder names → IDs internally (bypasses LLM ID corruption)
  const sourceResolved = await resolveMailFolderByName(cfg, userId, folderName);
  if (!sourceResolved.ok) return { isError: true, content: [{ type: "text", text: `Source: ${sourceResolved.error}` }] };

  const destResolved = await resolveMailFolderByName(cfg, userId, destinationName);
  if (!destResolved.ok) return { isError: true, content: [{ type: "text", text: `Destination: ${destResolved.error}` }] };

  const userPath = `/users/${encodeURIComponent(userId)}`;
  const movePath = `${userPath}/mailFolders${graphId(sourceResolved.folderId)}/move`;
  const result = await graphRequest<GraphMailFolder>(cfg, "POST", movePath, { destinationId: destResolved.folderId });

  if (!result.ok) {
    return { isError: true, content: [{ type: "text", text: result.error }] };
  }

  return {
    content: [{ type: "text", text: JSON.stringify({ moved: true, folderName, destination: destinationName, newDisplayName: result.data.displayName }, null, 2) }],
  };
}

/**
 * Create the Graph API tools for the A365 channel.
 *
 * Note: The execute functions use type assertions (e.g., `params as Parameters<...>`)
 * because TypeBox validates parameters at runtime against the schema before execution.
 * This provides runtime type safety while keeping the tool definitions concise.
 */
export function createGraphTools(cfg?: A365Config): AgentTool<TSchema, unknown>[] {
  const owner = cfg?.owner;
  const klipyKey = cfg?.klipyApiKey || process.env.KLIPY_API_KEY;

  const tools: AgentTool<TSchema, unknown>[] = [
    {
      name: "get_calendar_events",
      label: "Get Calendar Events",
      description: `Get calendar events for a user within a date range. ${owner ? `Default calendar owner: ${owner}` : "Requires userId parameter."}`,
      parameters: Type.Object({
        userId: Type.String({
          description: "User email or ID (use calendar owner's email for scheduling)",
        }),
        startDate: Type.String({
          description: "Start date/time in ISO format (e.g., 2024-01-15T00:00:00)",
        }),
        endDate: Type.String({
          description: "End date/time in ISO format (e.g., 2024-01-15T23:59:59)",
        }),
      }),
      execute: async (_toolCallId, params) => getCalendarEvents(cfg, params as Parameters<typeof getCalendarEvents>[1]),
    },
    {
      name: "create_calendar_event",
      label: "Create Calendar Event",
      description: `Create a new calendar event with optional attendees, location, and Teams meeting link. Use the attendees parameter to invite people. ${owner ? `Default calendar owner: ${owner}` : "Requires userId parameter."}`,
      parameters: Type.Object({
        userId: Type.String({
          description: "User email or ID whose calendar to create event on",
        }),
        subject: Type.String({ description: "Event subject/title" }),
        startDateTime: Type.String({
          description: "Start date/time in ISO format (e.g., 2024-01-15T14:00:00)",
        }),
        endDateTime: Type.String({
          description: "End date/time in ISO format (e.g., 2024-01-15T15:00:00)",
        }),
        timeZone: Type.Optional(
          Type.String({ description: "Time zone (default: from config or UTC)" }),
        ),
        attendees: Type.Optional(
          Type.Array(Type.String(), { description: "List of attendee email addresses" }),
        ),
        location: Type.Optional(Type.String({ description: "Meeting location" })),
        body: Type.Optional(Type.String({ description: "Event body/description" })),
        isOnlineMeeting: Type.Optional(
          Type.Boolean({ description: "Create as Teams meeting (default: false)" }),
        ),
      }),
      execute: async (_toolCallId, params) =>
        createCalendarEvent(cfg, params as Parameters<typeof createCalendarEvent>[1]),
    },
    {
      name: "update_calendar_event",
      label: "Update Calendar Event",
      description: "Update an existing calendar event. Can change subject, time, attendees, location, body, and Teams meeting status.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose calendar contains the event" }),
        eventId: Type.String({ description: "ID of the event to update" }),
        subject: Type.Optional(Type.String({ description: "New event subject/title" })),
        startDateTime: Type.Optional(
          Type.String({ description: "New start date/time in ISO format" }),
        ),
        endDateTime: Type.Optional(Type.String({ description: "New end date/time in ISO format" })),
        timeZone: Type.Optional(Type.String({ description: "Time zone for the new times" })),
        attendees: Type.Optional(
          Type.Array(Type.String(), { description: "Updated list of attendee email addresses (replaces existing attendees)" }),
        ),
        location: Type.Optional(Type.String({ description: "New meeting location" })),
        body: Type.Optional(Type.String({ description: "New event body/description" })),
        isOnlineMeeting: Type.Optional(
          Type.Boolean({ description: "Set to true to make it a Teams meeting" }),
        ),
      }),
      execute: async (_toolCallId, params) =>
        updateCalendarEvent(cfg, params as Parameters<typeof updateCalendarEvent>[1]),
    },
    {
      name: "delete_calendar_event",
      label: "Delete Calendar Event",
      description: "Delete a calendar event.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose calendar contains the event" }),
        eventId: Type.String({ description: "ID of the event to delete" }),
      }),
      execute: async (_toolCallId, params) =>
        deleteCalendarEvent(cfg, params as Parameters<typeof deleteCalendarEvent>[1]),
    },
    {
      name: "find_meeting_times",
      label: "Find Meeting Times",
      description:
        "Find available meeting times when all attendees are free. Uses Microsoft's scheduling assistant.",
      parameters: Type.Object({
        userId: Type.String({ description: "Organizer's email or ID" }),
        attendees: Type.Array(Type.String(), {
          description: "List of attendee email addresses",
        }),
        durationMinutes: Type.Number({ description: "Meeting duration in minutes" }),
        startDateTime: Type.String({
          description: "Search window start in ISO format",
        }),
        endDateTime: Type.String({
          description: "Search window end in ISO format",
        }),
        timeZone: Type.Optional(
          Type.String({ description: "Time zone (default: from config or UTC)" }),
        ),
      }),
      execute: async (_toolCallId, params) =>
        findMeetingTimes(cfg, params as Parameters<typeof findMeetingTimes>[1]),
    },
    {
      name: "send_email",
      label: "Send Email",
      description: "Send an email using Microsoft Graph.",
      parameters: Type.Object({
        userId: Type.String({ description: "Sender's email or ID (must have send permissions)" }),
        to: Type.Array(Type.String(), { description: "List of recipient email addresses" }),
        subject: Type.String({ description: "Email subject" }),
        body: Type.String({ description: "Email body content" }),
        cc: Type.Optional(Type.Array(Type.String(), { description: "CC recipients" })),
        bcc: Type.Optional(Type.Array(Type.String(), { description: "BCC recipients" })),
        importance: Type.Optional(
          Type.Union([Type.Literal("low"), Type.Literal("normal"), Type.Literal("high")], {
            description: "Email importance level",
          }),
        ),
      }),
      execute: async (_toolCallId, params) => sendEmail(cfg, params as Parameters<typeof sendEmail>[1]),
    },
    {
      name: "get_user_info",
      label: "Get User Info",
      description: "Get user profile information from Microsoft Graph.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID to look up" }),
      }),
      execute: async (_toolCallId, params) => getUserInfo(cfg, params as Parameters<typeof getUserInfo>[1]),
    },
    // --- Mail tools ---
    {
      name: "get_emails",
      label: "Get Emails",
      description: `List emails from a user's mailbox folder. ${owner ? `Default mailbox owner: ${owner}` : "Requires userId parameter."}`,
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose mailbox to read" }),
        folderName: Type.Optional(Type.String({ description: "Folder display name (default: 'inbox'). Examples: 'inbox', 'Drafts', 'Sent Items', 'Archive', or any custom folder name." })),
        top: Type.Optional(Type.Number({ description: "Number of emails to return (1-50, default: 10)" })),
        filter: Type.Optional(Type.String({ description: "OData filter (e.g., 'isRead eq false', 'hasAttachments eq true')" })),
        orderBy: Type.Optional(Type.String({ description: "Sort order (default: 'receivedDateTime desc')" })),
      }),
      execute: async (_toolCallId, params) => getEmails(cfg, params as Parameters<typeof getEmails>[1]),
    },
    {
      name: "read_email",
      label: "Read Email",
      description: "Read the full content of a specific email message by its ID.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose mailbox contains the message" }),
        messageId: Type.String({ description: "The message ID to read (from get_emails or search_emails)" }),
      }),
      execute: async (_toolCallId, params) => readEmail(cfg, params as Parameters<typeof readEmail>[1]),
    },
    {
      name: "search_emails",
      label: "Search Emails",
      description: "Search emails using KQL (Keyword Query Language). Examples: 'from:alice', 'subject:meeting', 'hasAttachments:true', 'received:2024-01-15'.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose mailbox to search" }),
        query: Type.String({ description: "KQL search query (e.g., 'from:alice subject:meeting')" }),
        top: Type.Optional(Type.Number({ description: "Number of results to return (1-50, default: 10)" })),
      }),
      execute: async (_toolCallId, params) => searchEmails(cfg, params as Parameters<typeof searchEmails>[1]),
    },
    {
      name: "move_email",
      label: "Move Email",
      description: "Move an email to a different folder. Use the folder's display name as destination.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose mailbox contains the message" }),
        messageId: Type.String({ description: "The message ID to move" }),
        destinationFolderName: Type.String({ description: "Display name of the target folder (e.g. 'Archive', 'Inbox', 'Projects')" }),
      }),
      execute: async (_toolCallId, params) => moveEmail(cfg, params as Parameters<typeof moveEmail>[1]),
    },
    {
      name: "delete_email",
      label: "Delete Email",
      description: "Delete an email message.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose mailbox contains the message" }),
        messageId: Type.String({ description: "The message ID to delete" }),
      }),
      execute: async (_toolCallId, params) => deleteEmail(cfg, params as Parameters<typeof deleteEmail>[1]),
    },
    {
      name: "mark_email_read",
      label: "Mark Email Read/Unread",
      description: "Mark an email as read or unread.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose mailbox contains the message" }),
        messageId: Type.String({ description: "The message ID to update" }),
        isRead: Type.Boolean({ description: "true to mark as read, false to mark as unread" }),
      }),
      execute: async (_toolCallId, params) => markEmailRead(cfg, params as Parameters<typeof markEmailRead>[1]),
    },
    {
      name: "get_mail_folders",
      label: "Get Mail Folders",
      description: "List mail folders in a user's mailbox. Returns top-level folders by default. Use parentFolderName to list child folders of a specific folder.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose mail folders to list" }),
        parentFolderName: Type.Optional(Type.String({ description: "Display name of parent folder to list child folders of (e.g. 'Inbox', 'Archive'). Omit for top-level folders." })),
      }),
      execute: async (_toolCallId, params) => getMailFolders(cfg, params as Parameters<typeof getMailFolders>[1]),
    },
    {
      name: "create_mail_folder",
      label: "Create Mail Folder",
      description: "Create a new mail folder in a user's mailbox. Can create top-level folders or subfolders.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose mailbox to create the folder in" }),
        displayName: Type.String({ description: "Name for the new folder" }),
        parentFolderName: Type.Optional(Type.String({ description: "Display name of parent folder to create a subfolder in (e.g. '_Legacy'). Omit for top-level folder." })),
      }),
      execute: async (_toolCallId, params) => createMailFolder(cfg, params as Parameters<typeof createMailFolder>[1]),
    },
    {
      name: "rename_mail_folder",
      label: "Rename Mail Folder",
      description: "Rename an existing mail folder. Use the folder's current display name to identify it.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID who owns the folder" }),
        folderName: Type.String({ description: "Current display name of the folder to rename (e.g. 'Old Projects')" }),
        newName: Type.String({ description: "New display name for the folder" }),
      }),
      execute: async (_toolCallId, params) => renameMailFolder(cfg, params as Parameters<typeof renameMailFolder>[1]),
    },
    {
      name: "delete_mail_folder",
      label: "Delete Mail Folder",
      description: "Delete a mail folder and all its contents. Use the folder's display name to identify it. Use with caution.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID who owns the folder" }),
        folderName: Type.String({ description: "Display name of the folder to delete (e.g. 'Old Archive')" }),
      }),
      execute: async (_toolCallId, params) => deleteMailFolder(cfg, params as Parameters<typeof deleteMailFolder>[1]),
    },
    {
      name: "move_mail_folder",
      label: "Move Mail Folder",
      description: "Move a mail folder into another folder. Use display names to identify folders (e.g., move 'Projects' into '_Legacy').",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID who owns the folder" }),
        folderName: Type.String({ description: "Display name of the folder to move (e.g. 'Projects')" }),
        destinationName: Type.String({ description: "Display name of the destination parent folder (e.g. '_Legacy')" }),
      }),
      execute: async (_toolCallId, params) => moveMailFolder(cfg, params as Parameters<typeof moveMailFolder>[1]),
    },
  ];

  // Add GIF tool only if Klipy API key is configured
  if (klipyKey) {
    tools.push({
      name: "send_gif",
      label: "Send GIF",
      description: "Search for and send an animated GIF inline in the conversation. The GIF is sent as a separate message. Use sparingly and only when it genuinely fits the moment.",
      parameters: Type.Object({
        query: Type.String({ description: "Search query describing the GIF to find (e.g. 'thumbs up', 'celebration', 'good morning')" }),
      }),
      execute: async (_toolCallId, params) => sendGif(cfg, params as { query: string }),
    });
  }

  return tools;
}
