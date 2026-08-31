/**
 * Model Context Protocol (MCP) Server for Microsoft Word
 * Cho phép các AI Agent ngoài (Claude Code, Cursor, Codex, OpenCode) tự động điều khiển Word qua stdio
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { tools, clientManager } from "./tools.js";

// Khởi tạo MCP Server
const server = new Server(
  {
    name: "word-bridge-mcp",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Trả về danh sách công cụ
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: tools.map((t) => ({
      name: t.name,
      description: t.description,
      inputSchema: t.inputSchema,
    })),
  };
});

// Xử lý thực thi công cụ
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;
  const tool = tools.find((t) => t.name === name);

  if (!tool) {
    return {
      isError: true,
      content: [
        {
          type: "text",
          text: `Không tìm thấy công cụ: ${name}`,
        },
      ],
    };
  }

  // Chuyển đổi tên tool từ 'word_xyz' sang kind 'xyz'
  const kind = name.startsWith("word_") ? name.slice(5) : name;
  const op = {
    kind,
    ...args,
  };

  try {
    const result = await clientManager.dispatchOp(op);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (err) {
    return {
      isError: true,
      content: [
        {
          type: "text",
          text: `Lỗi thực thi ${name}: ${err.message}${err.detail ? `\nChi tiết: ${JSON.stringify(err.detail)}` : ""}`,
        },
      ],
    };
  }
});

// Khởi chạy transport stdio
export async function startMcpServer() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("[MCP] Word Bridge MCP Server đã khởi chạy qua stdio.");
}

// Nếu chạy trực tiếp file này
if (process.argv[1]?.endsWith("mcp-server.js")) {
  startMcpServer().catch((err) => {
    console.error("[MCP] Lỗi khởi động:", err);
    process.exit(1);
  });
}
