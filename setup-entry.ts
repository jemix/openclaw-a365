import { defineSetupPluginEntry } from "openclaw/plugin-sdk/core";
import { a365Plugin } from "./src/channel.js";

export default defineSetupPluginEntry(a365Plugin);
