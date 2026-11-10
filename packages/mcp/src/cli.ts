import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createWordServer } from "./server.js";

// WORDINWEB_AUTHOR names the author of tracked changes and comments.
await createWordServer({ author: process.env.WORDINWEB_AUTHOR }).connect(new StdioServerTransport());
