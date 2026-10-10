import { ChatMessage, ProviderConfig, RetrievalChunk, ToolDefinition, ActionPlan, WordAction } from "./types.js";
import { logRequestStart, logResponseRaw, logResponseParsed, logError } from "./agent-logger.js";

type ChatCompletionResponse = {
  choices?: Array<{
    finish_reason?: string;
    message?: {
      content?: string;
      reasoning_content?: string;
      tool_calls?: Array<{
        id: string;
        type: "function";
        function: {
          name: string;
          arguments: string;
        };
      }>;
    };
  }>;
};

type StreamChunk = {
  choices?: Array<{
    finish_reason?: string;
    delta?: { content?: string; reasoning_content?: string; tool_calls?: any };
    text?: string;
    message?: { content?: string };
  }>;
};

type ModelsResponse = {
  data?: Array<{ id?: string }>;
};

export class LlmHttpError extends Error {
  status: number;

  details: string;

  constructor(status: number, details: string) {
    super(`LLM request failed (${status}): ${details}`);
    this.name = "LlmHttpError";
    this.status = status;
    this.details = details;
  }
}

// --- Word Tool Definitions ---

export const WORD_TOOLS: ToolDefinition[] = [
  // --- 感知类工具 (Perception) ---
  {
    type: "function",
    function: {
      name: "read_document",
      description: "动态读取文档中指定范围或类型的段落内容。用于在操作前确认目标位置、验证插入结果、或获取超出初始上下文的详细信息。建议在不确定段落索引或标题位置时先调用此工具确认。",
      parameters: {
        type: "object",
        properties: {
          mode: {
            type: "string",
            enum: ["paragraph_range", "heading_context", "selection", "cursor_surrounding"],
            description: "读取模式：paragraph_range=按段落范围读取；heading_context=读取指定标题及其下所有子内容直到下一个同级标题；selection=读取当前选区；cursor_surrounding=读取光标前后内容",
          },
          paragraph_index: { type: "number", description: "mode=paragraph_range 时的起始段落序号（从0开始）" },
          count: { type: "number", description: "mode=paragraph_range 时读取的段落数量，默认5" },
          heading_text: { type: "string", description: "mode=heading_context 时，要查找的标题文本。如果文档中存在多个同名标题（如目录和正文），工具会自动优先选择实际标题（Heading样式），并在返回结果中列出所有匹配项及其段落序号" },
          surrounding_chars: { type: "number", description: "mode=cursor_surrounding 时，光标前后各读取多少字符，默认500" },
        },
        required: ["mode"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_selection_info",
      description: "获取当前选区的精确信息，包括选区文本、起始/结束段落序号、是否仅为光标位置（无选中文本）。用于判断应使用 insert_at_cursor 还是 replace_selection。",
      parameters: {
        type: "object",
        properties: {},
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_document_stats",
      description: "获取文档统计信息：总段落数、总字符数、各级标题列表及其段落序号。用于快速了解文档整体结构。",
      parameters: {
        type: "object",
        properties: {},
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_paragraph_format",
      description: "获取指定段落的详细格式信息，包括段落内每个文本片段（run）的字体、字号、加粗、斜体、颜色。当需要精确了解段落内的内联格式变化时使用（例如：段落中部分文字使用了不同字体或加粗）。返回 paragraphFont（段落默认字体）和 runs 数组（每个 run 有独立字体信息）。如果 runs 只有 1 个，说明全段格式统一。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "要查看格式的段落序号（从0开始）" },
        },
        required: ["paragraph_index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_document_tables",
      description: "获取文档中所有 Word 表格的概览信息（不读取单元格内容）。返回每张表格的：表格序号（从0开始）、行数、列数、表格样式、起始段落序号、结束段落序号。用于：快速了解文档里有几张表格、定位表格在文档中的位置、决定下一步要读取哪张表的内容。**不要用 read_document 读取表格内的单元格文本**——表格单元格会被 Word 当作独立段落返回，LLM 看到的只是平铺的单元格文本，无法理解行列结构；本工具配合 read_table 才能高效获取结构化表格数据。",
      parameters: {
        type: "object",
        properties: {},
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_table",
      description: "读取指定表格的完整 2D 结构化内容（所有单元格的文本）。**这是读取表格数据的正确方式**——调用 get_document_tables 拿到 table_index 后再用本工具。返回内容按行/列格式化输出（每行形如 `[行N] 单元格1 | 单元格2 | ...`），便于 LLM 一次性理解整张表的结构。如果表格很大（>20 行），建议在结果中重点摘要用户关心的部分。",
      parameters: {
        type: "object",
        properties: {
          table_index: { type: "number", description: "表格序号（从0开始）。先调用 get_document_tables 确认索引范围。" },
        },
        required: ["table_index"],
      },
    },
  },

  // --- 操作类工具 (Action) ---
  {
    type: "function",
    function: {
      name: "insert_after_heading",
      description: "在指定标题后插入新内容。通过 heading_text 定位标题，支持指定格式。如果标题匹配失败，会插入到文档末尾。",
      parameters: {
        type: "object",
        properties: {
          heading_text: { type: "string", description: "要查找的标题文本" },
          content: { type: "string", description: "要插入的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式（如 <b>加粗</b>、<i>斜体</i>、<span style='font-family:黑体;font-size:14pt'>指定字体</span>）" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal。注意：content_format='html' 时此参数仍可用于设置段落样式",
          },
        },
        required: ["heading_text", "content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_at_cursor",
      description: "在当前光标位置插入内容。如果当前有选区，选区内容会被替换。这是最常用的插入工具，语义明确为「在光标处插入」。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "要插入的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式（如 <b>加粗</b>、<i>斜体</i>、<span style='font-family:黑体;font-size:14pt'>指定字体</span>）" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
          },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "replace_selection",
      description: "替换当前选中的文本内容。如果当前没有选区（仅光标），行为等同于 insert_at_cursor。建议先调用 get_selection_info 确认选区状态。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "替换后的新内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
          },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_at_end",
      description: "在文档末尾追加内容。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "要追加的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
          },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_at_start",
      description: "在文档开头插入内容。注意：空文档时可能需要特殊处理。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "要插入的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
          },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_after_paragraph",
      description: "在指定段落序号后插入内容。段落序号从文档结构中获取。注意：插入/删除操作后段落索引会变化，建议先调用 read_document 确认索引。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "目标段落的序号（从文档结构中获取，从0开始）" },
          content: { type: "string", description: "要插入的内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "段落级别样式，默认为 normal",
          },
        },
        required: ["paragraph_index", "content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_paragraph",
      description: "删除指定段落序号的内容。建议先调用 read_document 确认段落内容后再删除。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "要删除的段落序号（从0开始）" },
        },
        required: ["paragraph_index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_and_replace",
      description: "在文档中查找文本并替换为新文本。仅替换第一处匹配，大小写敏感。如需更多控制，请使用 find_and_replace_v2。",
      parameters: {
        type: "object",
        properties: {
          find_text: { type: "string", description: "要查找的文本" },
          replace_text: { type: "string", description: "替换后的文本" },
        },
        required: ["find_text", "replace_text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_and_replace_v2",
      description: "增强版查找替换。支持全文档替换、大小写控制、全词匹配，并返回替换次数。",
      parameters: {
        type: "object",
        properties: {
          find_text: { type: "string", description: "要查找的文本" },
          replace_text: { type: "string", description: "替换后的文本" },
          replace_all: { type: "boolean", description: "是否替换所有匹配项，默认 false（仅替换第一处）" },
          match_case: { type: "boolean", description: "是否区分大小写，默认 false" },
          match_whole_word: { type: "boolean", description: "是否全词匹配，默认 false" },
        },
        required: ["find_text", "replace_text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "apply_rich_format",
      description: "对指定段落或选区应用富文本格式（字体、颜色、加粗、斜体、超链接）。支持对段落内特定文字片段（通过 text_to_format 精确匹配）单独设置格式，实现段落内混合排版。",
      parameters: {
        type: "object",
        properties: {
          target_mode: {
            type: "string",
            enum: ["selection", "paragraph_index", "last_inserted"],
            description: "目标：selection=当前选区；paragraph_index=指定段落；last_inserted=上次插入的段落",
          },
          paragraph_index: { type: "number", description: "target_mode=paragraph_index 时的段落序号" },
          text_to_format: { type: "string", description: "段落内要格式化的精确文本片段。指定后只对该文本应用格式，而非整个段落/选区。例如段落中'重点内容'三个字需要加粗，则填'重点内容'" },
          font: {
            type: "object",
            properties: {
              name: { type: "string", description: "字体名称，如'微软雅黑'" },
              size: { type: "number", description: "字号（磅），如 14" },
              color: { type: "string", description: "字体颜色（十六进制），如 #FF0000" },
              bold: { type: "boolean", description: "是否加粗" },
              italic: { type: "boolean", description: "是否斜体" },
            },
          },
          hyperlink: {
            type: "object",
            properties: {
              text: { type: "string", description: "超链接显示文本" },
              url: { type: "string", description: "超链接URL" },
            },
          },
        },
        required: ["target_mode"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "replace_paragraph",
      description: "替换指定段落的全部文本内容。用于修改现有段落文字，保留段落格式。建议先调用 read_document 确认段落内容后再替换。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "要替换的段落序号（从0开始）" },
          content: { type: "string", description: "替换后的新内容。默认纯文本；当 content_format='html' 时可使用 HTML 标签实现内联格式" },
          content_format: { type: "string", enum: ["text", "html"], description: "内容格式：text=纯文本（默认），html=HTML格式（支持内联格式标签）" },
        },
        required: ["paragraph_index", "content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "set_paragraph_style",
      description: "修改指定段落的样式（如标题级别、列表格式）。用于调整已有段落的格式而不改变内容。",
      parameters: {
        type: "object",
        properties: {
          paragraph_index: { type: "number", description: "目标段落序号（从0开始）" },
          format: {
            type: "string",
            enum: ["normal", "heading1", "heading2", "heading3", "bullet_list", "numbered_list"],
            description: "要设置的格式",
          },
        },
        required: ["paragraph_index", "format"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "merge_paragraphs",
      description: "将两个相邻段落合并为一个段落。第二个段落的内容会追加到第一个段落的末尾。",
      parameters: {
        type: "object",
        properties: {
          first_paragraph_index: { type: "number", description: "第一个段落的序号（合并后内容保留在此位置）" },
          second_paragraph_index: { type: "number", description: "第二个段落的序号（合并后此段落会被删除）" },
          separator: { type: "string", description: "合并时两段之间的分隔符，默认为空格", default: " " },
        },
        required: ["first_paragraph_index", "second_paragraph_index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "reply_only",
      description: "仅回复文本，不执行任何文档操作。用于纯问答、解释、建议等场景。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "回复给用户的文本内容" },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "task_complete",
      description: "标记任务已完成，停止 Agent 循环。当所有操作都已执行完毕，或者任务不需要进一步操作时调用此工具。",
      parameters: {
        type: "object",
        properties: {
          summary: { type: "string", description: "任务完成的总结说明" },
        },
        required: ["summary"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "undo_last_action",
      description: "撤销最近的文档操作（类似 Ctrl+Z）。用于回退错误的插入、删除或替换。撤销后段落索引会变化，建议撤销后调用 read_document 确认当前文档状态。",
      parameters: {
        type: "object",
        properties: {
          count: { type: "number", description: "撤销步数，默认 1" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_table",
      description: "删除文档中指定的一张 Word 表格。先调用 get_document_tables 获取表格序号，再调用本工具删除。删除是不可逆的（但可用 undo_last_action 回退），建议删除前先调用 read_table 确认表格内容。",
      parameters: {
        type: "object",
        properties: {
          table_index: { type: "number", description: "要删除的表格序号（从0开始）。先调用 get_document_tables 确认索引。" },
        },
        required: ["table_index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_table",
      description:
        "在文档中插入一个真正的 Word 表格（不是 Markdown 表格语法）。当用户要求插入表格、对比表、列表矩阵等结构化内容时，必须使用本工具，而不是输出 Markdown 的 |---| 语法（Markdown 表格会原样插入到 Word 中无法阅读）。",
      parameters: {
        type: "object",
        properties: {
          location: {
            type: "string",
            enum: ["at_cursor", "at_end", "after_heading", "after_paragraph"],
            description: "插入位置。at_cursor=当前光标或选区处；at_end=文档末尾；after_heading/after_paragraph 需要配合 heading_text 或 paragraph_index。",
          },
          heading_text: {
            type: "string",
            description: "location=after_heading 时使用：要插入到其后的标题文本。",
          },
          paragraph_index: {
            type: "number",
            description: "location=after_paragraph 时使用：目标段落序号（从 0 开始）。",
          },
          headers: {
            type: "array",
            items: { type: "string" },
            description: "可选的表头行（第一行）。传空数组表示无表头，每行都是数据行。",
          },
          rows: {
            type: "array",
            items: {
              type: "array",
              items: { type: "string" },
            },
            description: "数据行，每个元素是一行（string[]）。短行会自动用空字符串补齐到与表头等宽。",
          },
          style: {
            type: "string",
            enum: ["TableGrid", "LightList", "MediumShading1", "NoBorders"],
            description: "表格样式。默认 TableGrid（带边框）。",
          },
        },
        required: ["location", "rows"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_mermaid_image",
      description: "用 mermaid 语法生成流程图/时序图/类图等并以图片插入 Word。支持 flowchart / sequenceDiagram / classDiagram / stateDiagram / erDiagram / gantt / pie / gitGraph / journey / C4 等所有 mermaid 语法。渲染失败时错误会原样回传，请修正语法后重试。",
      parameters: {
        type: "object",
        properties: {
          mermaid_code: {
            type: "string",
            description: "mermaid 源代码（以 `graph`/`flowchart`/`sequenceDiagram` 等关键字开头）。",
          },
          location: {
            type: "string",
            enum: ["at_cursor", "at_end", "after_heading", "after_paragraph"],
            description: "插入位置。at_cursor=当前光标或选区处（默认）；at_end=文档末尾；after_heading/after_paragraph 需配合 heading_text 或 paragraph_index。",
          },
          heading_text: {
            type: "string",
            description: "location=after_heading 时使用：要插入到其后的标题文本。",
          },
          paragraph_index: {
            type: "number",
            description: "location=after_paragraph 时使用：目标段落序号（从 0 开始）。",
          },
        },
        required: ["mermaid_code"],
      },
    },
  },
];

// --- Build document structure description ---

function describeFontBrief(font?: { name?: string; size?: number; color?: string; bold?: boolean; italic?: boolean }): string {
  if (!font) return "";
  const parts: string[] = [];
  if (font.name) parts.push(font.name);
  if (font.size) parts.push(`${font.size}pt`);
  if (font.bold) parts.push("加粗");
  if (font.italic) parts.push("斜体");
  if (font.color && !["#000000", "#000000ff", "#000"].includes(font.color.toLowerCase())) parts.push(font.color);
  return parts.length > 0 ? ` [${parts.join(" ")}]` : "";
}

export function describeDocumentStructure(structure: {
  totalParagraphs: number;
  totalCharacters?: number;
  paragraphs: Array<{ index: number; text: string; style: string; headingLevel?: number; isTable?: boolean; isList: boolean; charCount?: number; font?: { name?: string; size?: number; color?: string; bold?: boolean; italic?: boolean } }>;
  selection: { text: string; startParagraphIndex?: number; endParagraphIndex?: number };
}): string {
  const lines: string[] = [];
  lines.push(`文档共 ${structure.totalParagraphs} 段${structure.totalCharacters ? `，约 ${structure.totalCharacters} 字符` : ""}。结构概览：`);
  // Count tables so we can surface a hint up front.
  const tableParaCount = structure.paragraphs.filter((p) => p.isTable).length;
  if (tableParaCount > 0) {
    lines.push(`（注意：${tableParaCount} 个段落属于 Word 表格内的单元格。如需查看表格结构化数据，请使用 get_document_tables + read_table 工具，不要用 read_document 逐个读单元格文本。）`);
  }

  for (const p of structure.paragraphs) {
    const charInfo = p.charCount ? `(${p.charCount}字)` : "";
    const fontInfo = describeFontBrief(p.font);
    const tableMark = p.isTable ? "[TBL] " : "";
    if (p.headingLevel) {
      lines.push(`  ${"#".repeat(p.headingLevel)} [段落${p.index}]${charInfo}${fontInfo} ${tableMark}${p.text}`);
    } else if (p.isList) {
      lines.push(`  - [段落${p.index}]${charInfo}${fontInfo} ${tableMark}${p.text}`);
    } else {
      lines.push(`  [段落${p.index}]${charInfo}${fontInfo} ${tableMark}${p.text.slice(0, 100)}`);
    }
  }

  if (structure.selection.text) {
    const selInfo = structure.selection.startParagraphIndex !== undefined
      ? `（位于段落 ${structure.selection.startParagraphIndex}${structure.selection.endParagraphIndex !== structure.selection.startParagraphIndex ? ` ~ ${structure.selection.endParagraphIndex}` : ""}）`
      : "";
    lines.push(`\n当前选区${selInfo}：${structure.selection.text.slice(0, 500)}`);
  } else {
    lines.push(`\n当前无选中文本（光标模式）`);
  }

  lines.push(`\n注意：以上字体为段落级别默认字体。Word 支持段落内内联格式变化（如部分文字使用不同字体、加粗等），如需查看某段落的详细格式分布，请使用 get_paragraph_format 工具。`);

  return lines.join("\n");
}

// --- Build system prompt based on mode and document structure ---

function buildSystemPrompt(
  hasTools: boolean,
  insertMode: string
): ChatMessage {
  let content = `你是一个面向 Word 文档编辑的智能 Agent。你的工作是理解用户意图，通过「感知 → 思考 → 操作」的循环来完成任务。

## 核心规则
1. 根据文档结构和用户意图选择最合适的工具和参数
2. 如果用户要求在某个位置插入内容，优先使用 insert_after_heading 或 insert_after_paragraph，而不是让用户手动定位
3. 如果用户要求删除内容，使用 delete_paragraph 或 find_and_replace_v2
4. 如果用户只是提问或需要建议，使用 reply_only
5. 当使用 insert/replace 工具时，content 参数应该是可直接写入 Word 的纯文本，不要使用 Markdown 标记
6. format 参数用于指定插入内容的格式，默认为 normal
7. 在引用知识库时标注来源编号

## ⚠️ 目标完成准则（最重要）
**你必须严格区分「打算做」和「已经完成」：**
- 只有当你通过工具返回的结果**确认**所有操作都成功执行，并且**验证**了结果符合用户预期时，才能调用 task_complete
- 如果用户的请求包含多个部分（如"先做A，再做B"），你必须**完成所有部分**后才能结束
- 如果用户的请求是"修改/替换/删除所有X"，你必须**确认所有X都被处理**后才能结束
- **禁止在未验证的情况下假设操作成功** — 每次操作后都应读取文档确认
- 如果不确定是否完成，**继续执行**而不是调用 task_complete

**过早结束是严重错误** — 宁可多验证一次，也不要提前结束任务

## 感知优先原则
- 如果不确定文档当前状态、目标位置、或索引是否有效，先调用 read_document 或 get_selection_info 确认
- 插入/删除操作后段落索引会变化，后续操作应基于新的索引
- 如果 find_and_replace_v2 返回 replaced=0，说明查找失败，不要继续基于该假设执行后续操作
- content 参数必须是非空字符串
- 注意：段落级别字体信息只是该段落的默认/主字体。Word 文档支持在同一段落内对不同文字片段设置不同格式（内联格式），如需精确了解段落内的格式分布（例如某段落中部分文字加粗、部分使用不同字体），请使用 get_paragraph_format 工具获取逐片段的格式详情
- 涉及 Word 表格的读取时**必须**使用 get_document_tables + read_table；不要用 read_document 读表格里的单元格文本——单元格会被 Word 当作独立段落平铺返回，LLM 无法看到行列结构`;

  if (hasTools) {
    content += "\n\n## 可用工具清单\n\n你有以下工具可以调用。每个工具都有明确的适用场景，请根据用户意图选择最合适的工具。\n";
    content += "\n---\n\n### 感知类工具 — 先看再动，了解文档现状\n\n";
    content += "#### 'read_document'\n";
    content += "读取文档中指定范围的内容。**在不确定段落索引、标题位置、或需要验证操作结果时，必须先调用此工具。**\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| mode | string | 是 | 读取模式，可选值见下方 |\n";
    content += "| paragraph_index | number | 条件 | mode='paragraph_range' 时，起始段落序号（从 0 开始） |\n";
    content += "| count | number | 否 | mode='paragraph_range' 时，读取的段落数量，默认 5 |\n";
    content += "| heading_text | string | 条件 | mode='heading_context' 时，要查找的标题文本 |\n";
    content += "| surrounding_chars | number | 否 | mode='cursor_surrounding' 时，光标前后各读取多少字符，默认 500 |\n\n";
    content += "mode 可选值：\n";
    content += "- 'paragraph_range' — 从 paragraph_index 开始连续读取 count 段\n";
    content += "- 'heading_context' — 读取 heading_text 对应标题及其下属所有内容（直到遇到同级标题）\n";
    content += "- 'selection' — 读取当前选中的内容\n";
    content += "- 'cursor_surrounding' — 读取光标前后的文本\n\n";
    content += "使用场景：插入前确认目标位置 | 操作后验证结果 | 查看标题下的完整内容 | 获取超出初始上下文的段落\n\n";
    content += "示例：read_document({ mode: 'heading_context', heading_text: '第三章' }) → 读取'第三章'标题下的所有内容\n\n";
    content += "---\n\n";
    content += "#### 'get_selection_info'\n";
    content += "获取当前选区的精确信息。**无参数**，返回选区文本、起始/结束段落序号、是否仅为光标（无选中文本）。\n\n";
    content += "使用场景：判断应该用 insert_at_cursor 还是 replace_selection | 确认用户当前操作意图\n\n";
    content += "---\n\n";
    content += "#### 'get_document_stats'\n";
    content += "获取文档统计信息。**无参数**，返回总段落数、总字符数、各级标题列表及其段落序号。\n\n";
    content += "使用场景：快速了解文档整体结构 | 确定标题是否存在 | 评估文档规模\n\n";
    content += "---\n\n";
    content += "#### 'get_paragraph_format'\n";
    content += "获取指定段落的**内联格式详情**。返回段落默认字体 + 每个文字片段（run）的独立格式信息。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| paragraph_index | number | 是 | 要查看的段落序号（从 0 开始） |\n\n";
    content += "返回值解读：\n- runs 只有 1 个 → 全段格式统一\n- runs 有多个 → 段落内存在格式变化（部分加粗、不同字体等）\n\n";
    content += "使用场景：需要精确了解段落内格式分布 | 判断是否需要修改内联格式\n\n";
    content += "---\n\n";
    content += "#### 'get_document_tables'\n";
    content += "获取文档中所有 Word 表格的**概览信息**（不读取单元格内容）。**无参数**，返回每张表格的：表格序号、行数、列数、表格样式、起始段落序号、结束段落序号。\n\n";
    content += "**重要**：不要用 read_document 读取表格内的单元格文本——Word 会把每个单元格当成独立段落返回，LLM 看到的只是平铺的单元格文本，无法理解行列结构。读取表格请用本工具 + read_table。\n\n";
    content += "使用场景：发现文档里有几张表格 | 定位表格在文档中的位置 | 决定下一步读哪张表\n\n";
    content += "---\n\n";
    content += "#### 'read_table'\n";
    content += "读取指定表格的**完整 2D 结构化内容**（所有单元格的文本）。这是读取表格数据的正确方式。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| table_index | number | 是 | 表格序号（从 0 开始）。先调用 get_document_tables 确认索引 |\n\n";
    content += "返回格式：每行形如 `[行N] 单元格1 | 单元格2 | ...`。\n\n";
    content += "使用场景：查看整张表的内容 | 摘要表格数据 | 基于表格内容做修改决策\n\n";
    content += "---\n\n";
    content += "#### 'delete_table'\n";
    content += "删除文档中指定的一张 Word 表格。删除不可逆（但可用 undo_last_action 回退），**建议删除前先用 read_table 确认表格内容**。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| table_index | number | 是 | 要删除的表格序号（从 0 开始）。先调用 get_document_tables 确认 |\n\n";
    content += "使用场景：删除废弃/重复表格 | 按指令清理文档中的表格\n\n";
    content += "---\n\n";
    content += "### 操作类工具 — 安全可自动迭代\n\n";
    content += "> 以下工具可在 Agent 循环中自动执行（无需用户确认），建议操作后用 read_document 验证结果。\n\n";
    content += "#### 'insert_at_cursor' [最常用]\n";
    content += "在**光标位置**插入内容。如果有选区，选区内容会被替换。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| content | string | 是 | 要插入的文本内容 |\n";
    content += "| content_format | enum | 否 | 'text'（默认）或 'html'（支持内联格式标签） |\n";
    content += "| format | enum | 否 | 段落样式：'normal'（默认）、'heading1'-'heading3'、'bullet_list'、'numbered_list' |\n\n";
    content += "示例：insert_at_cursor({ content: '这是新增的段落', format: 'heading2' })\n\n";
    content += "---\n\n";
    content += "#### 'insert_after_heading'\n";
    content += "在指定**标题后面**插入内容。通过标题文本定位，匹配失败则插入到文档末尾。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| heading_text | string | 是 | 要定位的标题文本（精确匹配） |\n";
    content += "| content | string | 是 | 要插入的内容 |\n";
    content += "| content_format | enum | 否 | 'text'（默认）或 'html' |\n";
    content += "| format | enum | 否 | 段落样式，默认 normal |\n\n";
    content += "示例：insert_after_heading({ heading_text: '第三章', content: '这是新增的内容' })\n\n";
    content += "---\n\n";
    content += "#### 'insert_after_paragraph'\n";
    content += "在指定**段落序号后**插入内容。段落序号从文档结构中获取。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| paragraph_index | number | 是 | 目标段落的序号（从文档结构中获取，从 0 开始） |\n";
    content += "| content | string | 是 | 要插入的内容 |\n";
    content += "| content_format | enum | 否 | 'text'（默认）或 'html' |\n";
    content += "| format | enum | 否 | 段落样式，默认 normal |\n\n";
    content += "**注意**：插入/删除后段落索引会变化，必须先 read_document 确认索引。\n\n";
    content += "---\n\n";
    content += "#### 'delete_paragraph'\n";
    content += "删除指定段落序号的内容。**不可撤销**（但可用 undo_last_action 回退）。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| paragraph_index | number | 是 | 要删除的段落序号（从 0 开始） |\n\n";
    content += "**注意**：删除后段落索引会变化，必须先 read_document 确认索引。\n\n";
    content += "---\n\n";
    content += "#### 'replace_selection'\n";
    content += "替换当前**选中文本**。空选区时等同于 insert_at_cursor。建议先用 get_selection_info 确认状态。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| content | string | 是 | 替换后的新内容 |\n";
    content += "| content_format | enum | 否 | 'text'（默认）或 'html' |\n";
    content += "| format | enum | 否 | 段落样式，默认 normal |\n\n";
    content += "---\n\n";
    content += "#### 'replace_paragraph'\n";
    content += "替换指定段落的**全部文本**，保留原有格式。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| paragraph_index | number | 是 | 目标段落序号（从 0 开始） |\n";
    content += "| content | string | 是 | 替换后的新内容 |\n";
    content += "| content_format | enum | 否 | 'text'（默认）或 'html' |\n\n";
    content += "---\n\n";
    content += "#### 'insert_at_end'\n";
    content += "在文档**末尾追加**内容。参数同 insert_at_cursor。\n\n";
    content += "---\n\n";
    content += "#### 'find_and_replace_v2' [推荐用于查找替换]\n";
    content += "增强版全文查找替换，**返回替换了多少处**。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| find_text | string | 是 | 要查找的文本 |\n";
    content += "| replace_text | string | 是 | 替换为的文本 |\n";
    content += "| replace_all | boolean | 否 | 是否替换全部匹配项，默认 false（仅第一处） |\n";
    content += "| match_case | boolean | 否 | 是否区分大小写，默认 false |\n";
    content += "| match_whole_word | boolean | 否 | 是否全词匹配，默认 false |\n\n";
    content += "**注意**：如果返回 replaced=0，说明没找到，不要再基于假设继续操作。\n\n";
    content += "示例：find_and_replace_v2({ find_text: '旧名称', replace_text: '新名称', replace_all: true }) → 将所有'旧名称'替换为'新名称'\n\n";
    content += "---\n\n";
    content += "#### 'set_paragraph_style'\n";
    content += "修改指定段落的**样式**（标题级别、列表格式），不改变内容。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| paragraph_index | number | 是 | 目标段落序号 |\n";
    content += "| format | enum | 是 | 'normal'、'heading1'-'heading3'、'bullet_list'、'numbered_list' |\n\n";
    content += "---\n\n";
    content += "#### 'apply_rich_format'\n";
    content += "对段落或选区应用**富文本格式**（字体、颜色、加粗、斜体、超链接）。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| target_mode | enum | 是 | 'selection'=当前选区；'paragraph_index'=指定段落；'last_inserted'=上次插入的段落 |\n";
    content += "| paragraph_index | number | 条件 | target_mode='paragraph_index' 时必填 |\n";
    content += "| text_to_format | string | 否 | 段落内要格式化的精确文本片段（留空=格式化整段） |\n";
    content += "| font.name | string | 否 | 字体名称，如 '微软雅黑' |\n";
    content += "| font.size | number | 否 | 字号（磅），如 14 |\n";
    content += "| font.color | string | 否 | 颜色，如 '#FF0000' |\n";
    content += "| font.bold | boolean | 否 | 是否加粗 |\n";
    content += "| font.italic | boolean | 否 | 是否斜体 |\n";
    content += "| hyperlink.text | string | 否 | 超链接显示文本 |\n";
    content += "| hyperlink.url | string | 否 | 超链接 URL |\n\n";
    content += "示例：apply_rich_format({ target_mode: 'paragraph_index', paragraph_index: 3, text_to_format: '重点', font: { bold: true, color: '#FF0000' } }) → 把第3段中'重点'二字设为红色加粗\n\n";
    content += "---\n\n";
    content += "#### 'merge_paragraphs'\n";
    content += "合并两个**相邻段落**。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| first_paragraph_index | number | 是 | 保留段落的序号 |\n";
    content += "| second_paragraph_index | number | 是 | 被合并段落的序号（合并后删除） |\n";
    content += "| separator | string | 否 | 两段之间的分隔符，默认空格 |\n\n";
    content += "---\n\n";
    content += "#### 'undo_last_action'\n";
    content += "撤销最近的文档操作（类似 Ctrl+Z）。用于回退错误的插入、删除或替换。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| count | number | 否 | 撤销步数，默认 1 |\n\n";
    content += "**注意**：撤销后段落索引会变化，必须调用 read_document 确认当前状态。\n\n";
    content += "---\n\n";
    content += "### 高风险操作 — 需要用户确认后才执行\n\n";
    content += "> 以下工具不会在 Agent 循环中自动执行，会先展示给用户确认。\n\n";
    content += "#### 'insert_at_start'\n在文档**开头**插入内容。空文档时有风险。\n\n";
    content += "#### 'find_and_replace'\n简单版查找替换（只能替换第一处，大小写敏感）。推荐使用 find_and_replace_v2 代替。\n\n";
    content += "---\n\n";
    content += "### 控制类工具\n\n";
    content += "#### 'reply_only'\n仅回复文本，**不操作文档**。用于纯问答、解释、建议等不需要编辑的场景。\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| content | string | 是 | 回复给用户的文本 |\n\n";
    content += "---\n\n";
    content += "#### 'task_complete' [必须调用]\n";
    content += "**标记任务完成**，结束 Agent 循环。\n\n";
    content += "**⚠️ 调用条件（必须同时满足）：**\n";
    content += "1. 用户请求的**所有操作**都已执行完成\n";
    content += "2. 通过 read_document 或其他感知工具**验证**了操作结果\n";
    content += "3. 结果**符合用户预期**（如替换了所有目标、插入了完整内容等）\n\n";
    content += "**禁止调用的情况：**\n";
    content += "- 只完成了用户请求的一部分（如用户要求处理5个章节，只处理了3个）\n";
    content += "- 操作后未验证结果\n";
    content += "- 不确定是否所有目标都已达成\n\n";
    content += "| 参数 | 类型 | 必填 | 说明 |\n|------|------|------|------|\n";
    content += "| summary | string | 是 | 任务完成的总结说明（会展示给用户） |\n\n";
    content += "---\n\n";
    content += "## 通用规范\n\n";
    content += "### HTML 内联格式（content_format='html'）\n";
    content += "所有插入/替换工具都支持 content_format='html'，可直接在 content 中使用 HTML 标签：\n";
    content += "- <b>加粗</b>、<i>斜体</i>、<u>下划线</u>\n";
    content += "- <span style='font-family:黑体;font-size:14pt'>指定字体</span>\n";
    content += "- 混合示例：content_format='html', content='这是<b>重点</b>和<i>斜体</i>的混合段落'\n\n";
    content += "### format 参数枚举\n";
    content += "'normal'（默认）| 'heading1' | 'heading2' | 'heading3' | 'bullet_list' | 'numbered_list'\n\n";
    content += "### content 参数注意事项\n";
    content += "- 必须是**非空字符串**，空字符串会导致错误\n";
    content += "- 不要使用 Markdown 标记（如 **、#、代码块 等），应使用 HTML 标签或直接纯文本\n\n";
    content += "---\n\n";
    content += "## 自我迭代工作流\n\n";
    content += "你可以在一个 Agent 循环中完成 感知 → 操作 → 验证 → 结束 的完整流程：\n\n";
    content += "| 步骤 | 做什么 | 用什么工具 |\n|------|--------|-----------|\n";
    content += "| 1. 感知 | 了解文档现状 | read_document、get_document_stats、get_selection_info、get_document_tables、read_table |\n";
    content += "| 2. 操作 | 执行修改 | insert_at_cursor、find_and_replace_v2、replace_paragraph 等 |\n";
    content += "| 3. 验证 | 确认修改结果 | read_document、get_selection_info、read_table |\n";
    content += "| 4. 修正 | 不满足预期则继续调整 | 回到步骤 2 |\n";
    content += "| 5. 结束 | **验证通过后**标记任务完成 | task_complete |\n\n";
    content += "**示例**：用户说「把所有'旧名称'换成'新名称'」\n";
    content += "1. find_and_replace_v2({ find_text: '旧名称', replace_text: '新名称', replace_all: true }) → 收到 replaced=3\n";
    content += "2. read_document({ mode: 'paragraph_range', paragraph_index: 0, count: 10 }) → 确认替换后的内容\n";
    content += "3. task_complete({ summary: '已将文档中所有 3 处旧名称替换为新名称' }) → 结束\n\n";
    content += "**示例**：用户说「在第3章后插入答案」\n";
    content += "1. read_document({ mode: 'heading_context', heading_text: '第三章' }) → 确认目标位置和段落索引\n";
    content += "2. insert_after_paragraph({ paragraph_index: 42, content: '答案内容...' }) → 插入\n";
    content += "3. read_document({ mode: 'paragraph_range', paragraph_index: 40, count: 10 }) → **必须读取**，验证插入结果+获取新索引\n";
    content += "4. task_complete({ summary: '已在第3章后插入答案' }) → 结束\n\n";
    content += "**重要提示**：\n";
    content += "- 不要把'打算做'和'已经做'混淆——只有工具返回的结果才能确认操作已成功\n";
    content += "- 不确定时优先 read_document，不要猜\n";
    content += "- 如果 find_and_replace_v2 返回 replaced=0，说明查找失败，检查 find_text 是否正确\n";
    content += "- **严禁重复操作**：如果工具返回结果显示操作已成功（如「已在段落 X 后插入内容」），说明该操作已完成，绝对不要对同一目标再次执行相同操作。直接进行下一步或调用 task_complete\n";
    content += "- **多题场景**：用户要求回答多道题目时，每道题插入完成后不要回头重做已完成的题目，依次处理即可\n";
    content += "- **段落索引失效规则（最重要）**：insert_after_paragraph 和 delete_paragraph 会改变文档中后续段落的索引号。因此：\n";
    content += "  - 每次 insert_after_paragraph 或 delete_paragraph 执行成功后，**必须先调用 read_document 获取最新段落索引**，然后才能进行下一次插入/删除/替换操作\n";
    content += "  - **绝对禁止**连续调用 insert_after_paragraph 或 delete_paragraph 而中间不插入 read_document\n";
    content += "  - 正确流程：read_document → 操作 → read_document（验证+获取新索引）→ 下一个操作\n";
    content += "  - 同样，undo_last_action 也会改变段落索引，撤销后必须 read_document\n";
    content += "- **目标检查清单**：在调用 task_complete 前，回顾用户原始请求，确认：\n";
    content += "  □ 是否完成了用户要求的所有操作？\n";
    content += "  □ 是否验证了操作结果？\n";
    content += "  □ 是否有遗漏的目标未处理？\n";
    content += "  如果任何一项为「否」，继续执行而不是调用 task_complete\n";
  }

  if (insertMode === "chat_only") {
    content += `

当前模式为「仅对话」，请只使用 reply_only 工具回复用户，不要执行任何文档操作。`;
  } else if (insertMode === "replace_selection") {
    content += `

当前模式为「替换选区」，请优先使用 replace_selection 或 insert_at_cursor 工具。`;
  } else if (insertMode === "append_end") {
    content += `

当前模式为「追加到文末」，请优先使用 insert_at_end 工具。`;
  } else if (insertMode === "smart_action") {
    content += `

当前模式为「智能操作」，请根据用户意图自主选择最合适的工具。建议在不确定时先调用感知类工具确认文档状态。`;
  }

  // NOTE: documentStructureDescription is intentionally NOT appended here.
  // It is high-frequency changing content (cursor moves, edits) and would
  // break the LLM provider's prefix-based prompt cache (DeepSeek context
  // caching, llama.cpp KV cache). It is injected as a trailing system
  // message in buildPayload() instead, keeping this system prompt as a
  // stable prefix that can be cached across requests.

  return { role: "system", content };
}

// --- Build payload ---

/**
 * Build the final messages array sent to the LLM.
 *
 * Message ordering is optimized for prefix-based prompt caching
 * (DeepSeek context caching, llama.cpp KV cache, Anthropic prompt cache):
 *
 *   [1] system: stable tool instructions + core rules (buildSystemPrompt)
 *   [2] ...conversation history (user / assistant / tool)...
 *   [3] system: dynamic context (retrieved chunks + document structure +
 *               document context + selection) — high-frequency changing,
 *               placed at the END so it does not break the cached prefix.
 *
 * Previously, retrieved chunks were the 2nd message and document structure
 * was appended to the system prompt, which made the stable prefix length ~0
 * and cache hit rate near zero.
 */
function buildPayload(
  config: ProviderConfig,
  messages: ChatMessage[],
  contextChunks: RetrievalChunk[],
  stream: boolean,
  tools?: ToolDefinition[],
  insertMode?: string,
  documentStructureDescription?: string,
  dynamicContext?: { documentContext?: string; selection?: string }
): Record<string, any> {
  const contextText = contextChunks.map((chunk, idx) => `[${idx + 1}] ${chunk.fileName}: ${chunk.text}`).join("\n\n");

  const hasTools = !!(tools && tools.length > 0);
  // Stable system prompt — no dynamic content appended.
  const systemPrompt = buildSystemPrompt(hasTools, insertMode || "smart_action");

  // Sanitize messages: remove orphaned tool_calls that have no matching tool
  // response messages.  This prevents 400 errors from the LLM API when a
  // session contains assistant messages with tool_calls that were never
  // followed up (e.g. user abandoned an action_plan, or an error occurred
  // before tool results could be sent).
  const sanitized = sanitizeMessages(messages);

  // Assemble the trailing dynamic-context system message.
  // All high-frequency changing content goes here, at the END of the array,
  // so the stable system prompt + conversation history prefix remains cacheable.
  const dynamicParts: string[] = [];
  if (contextText) {
    dynamicParts.push(`以下是可用知识片段：\n${contextText}\n\n请尽可能基于这些片段回答，并标注来源编号。`);
  } else {
    dynamicParts.push("当前没有检索到知识库片段，你可以基于用户输入给出通用建议。");
  }
  if (documentStructureDescription) {
    dynamicParts.push(`当前文档结构：\n${documentStructureDescription}`);
  }
  if (dynamicContext?.documentContext) {
    dynamicParts.push(`文档上下文:\n${dynamicContext.documentContext}`);
  }
  if (dynamicContext?.selection) {
    dynamicParts.push(`当前选区:\n${dynamicContext.selection}`);
  }
  const trailingSystem: ChatMessage = {
    role: "system",
    content: dynamicParts.join("\n\n"),
  };

  const payload: Record<string, any> = {
    model: config.model,
    messages: [systemPrompt, ...sanitized, trailingSystem],
    temperature: config.temperature ?? 0.2,
    max_tokens: config.maxTokens ?? 4096,
    stream,
  };

  if (hasTools) {
    payload.tools = tools;
    payload.tool_choice = "auto";
  }

  // ── Thinking / reasoning mode ──────────────────────────────────────────
  if (config.enableThinking) {
    const fmt = config.thinkingFormat ?? "deepseek";
    if (fmt === "deepseek") {
      // DeepSeek: thinking passed via extra_body (top-level JSON field)
      payload.thinking = { type: "enabled" };
      payload.reasoning_effort = config.thinkingEffort ?? "high";
    } else {
      // OpenAI-compatible / Agnes: chat_template_kwargs.enable_thinking
      payload.chat_template_kwargs = { enable_thinking: true };
    }
  }

  return payload;
}

/**
 * Walk through the message array and fix three kinds of inconsistencies:
 *
 * 1. Assistant messages with `tool_calls` where some tool_call IDs have no
 *    matching `tool` role response → strip `tool_calls`, keep text content.
 * 2. `tool` role messages whose `tool_call_id` does not belong to any
 *    preceding assistant `tool_calls` → drop the orphaned tool message.
 * 3. Duplicate `tool` messages with the same `tool_call_id` → keep only the
 *    first occurrence (APIs require exactly one tool response per tool_call).
 *
 * This prevents 400 errors from OpenAI-compatible APIs that enforce both:
 *   - "every tool_call must be followed by a tool response"
 *   - "every tool message must be a response to a preceding tool_call"
 */
function sanitizeMessages(messages: ChatMessage[]): ChatMessage[] {
  // Pass 1: identify which assistant tool_call IDs have matching tool responses
  // and which tool messages are orphaned.
  const validToolCallIds = new Set<string>();
  const assistantIdxWithToolCalls: Array<{ index: number; ids: Set<string> }> = [];

  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    if (msg.role === "assistant" && msg.tool_calls && msg.tool_calls.length > 0) {
      const ids = new Set(msg.tool_calls.map((tc) => tc.id));
      assistantIdxWithToolCalls.push({ index: i, ids });
    }
  }

  // For each assistant with tool_calls, check if all IDs have responses
  for (const entry of assistantIdxWithToolCalls) {
    const remaining = new Set(entry.ids);
    let j = entry.index + 1;
    while (j < messages.length && messages[j].role === "tool") {
      const tcId = messages[j].tool_call_id;
      if (tcId) {
        remaining.delete(tcId);
      }
      j++;
    }
    if (remaining.size === 0) {
      // All tool_calls have responses — mark them as valid
      for (const id of entry.ids) {
        validToolCallIds.add(id);
      }
    }
  }

  // Pass 2: build sanitized result, deduplicating tool messages by tool_call_id
  const result: ChatMessage[] = [];
  const seenToolCallIds = new Set<string>();
  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];

    if (msg.role === "assistant" && msg.tool_calls && msg.tool_calls.length > 0) {
      const ids = assistantIdxWithToolCalls.find((e) => e.index === i)?.ids;
      if (ids && [...ids].every((id) => validToolCallIds.has(id))) {
        // All tool_calls are valid — keep as-is
        result.push(msg);
      } else {
        // Orphaned tool_calls — strip them, keep text and reasoning only
        result.push({ role: "assistant", content: msg.content || "", reasoning_content: msg.reasoning_content });
      }
    } else if (msg.role === "tool" && msg.tool_call_id) {
      if (validToolCallIds.has(msg.tool_call_id) && !seenToolCallIds.has(msg.tool_call_id)) {
        // Valid and first occurrence — keep it
        seenToolCallIds.add(msg.tool_call_id);
        result.push(msg);
      }
      // else: orphaned or duplicate tool message — drop it
    } else {
      result.push(msg);
    }
  }

  return result;
}

// --- Parse action plan from LLM response ---

export function parseActionPlanFromToolCalls(
  toolCalls: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }>
): ActionPlan {
  const actions: WordAction[] = [];

  for (const tc of toolCalls) {
    let params: Record<string, any> = {};
    try {
      params = JSON.parse(tc.function.arguments);
    } catch {
      params = { raw_arguments: tc.function.arguments };
    }

    // Generate human-readable description
    const desc = describeAction(tc.function.name, params);

    actions.push({
      action: tc.function.name,
      params,
      description: desc,
      toolCallId: tc.id,
    });
  }

  return {
    actions,
    explanation: actions.map((a) => a.description).join("；"),
  };
}

function describeAction(actionName: string, params: Record<string, any>): string {
  const fmtSuffix = params.format && params.format !== "normal" ? `（格式: ${params.format}）` : "";
  switch (actionName) {
    // 感知类工具
    case "read_document":
      if (params.mode === "paragraph_range") return `读取段落 ${params.paragraph_index ?? 0} 起共 ${params.count ?? 5} 段`;
      if (params.mode === "heading_context") return `读取标题"${params.heading_text}"及其子内容`;
      if (params.mode === "selection") return "读取当前选区内容";
      if (params.mode === "cursor_surrounding") return `读取光标周围 ${params.surrounding_chars ?? 500} 字符`;
      return `读取文档（模式: ${params.mode}）`;
    case "get_selection_info":
      return "获取当前选区信息";
    case "get_document_stats":
      return "获取文档统计信息";
    case "get_paragraph_format":
      return `获取段落${params.paragraph_index}的详细格式`;
    // 操作类工具
    case "insert_after_heading":
      return `在标题"${params.heading_text}"后插入${params.content_format === "html" ? "HTML格式" : ""}内容${fmtSuffix}`;
    case "insert_at_cursor":
      return `在光标处插入${params.content_format === "html" ? "HTML格式" : ""}内容${fmtSuffix}`;
    case "replace_selection":
      return `替换选区${params.content_format === "html" ? "HTML格式" : ""}内容${fmtSuffix}`;
    case "insert_at_end":
      return `追加${params.content_format === "html" ? "HTML格式" : ""}内容到文档末尾${fmtSuffix}`;
    case "insert_at_start":
      return `插入${params.content_format === "html" ? "HTML格式" : ""}内容到文档开头${fmtSuffix}`;
    case "insert_after_paragraph":
      return `在段落${params.paragraph_index}后插入${params.content_format === "html" ? "HTML格式" : ""}内容${fmtSuffix}`;
    case "delete_paragraph":
      return `删除段落${params.paragraph_index}`;
    case "find_and_replace":
      return `将"${params.find_text}"替换为"${params.replace_text}"`;
    case "find_and_replace_v2":
      return `查找替换"${params.find_text}"→"${params.replace_text}"${params.replace_all ? "（全部）" : "（首处）"}${params.match_case ? " 区分大小写" : ""}`;
    case "apply_rich_format":
      if (params.hyperlink) return `对${params.target_mode}应用超链接"${params.hyperlink.text}"`;
      if (params.text_to_format) return `对段落内文本"${params.text_to_format}"应用富文本格式`;
      return `对${params.target_mode}应用富文本格式`;
    case "replace_paragraph":
      return `替换段落${params.paragraph_index}的内容${params.content_format === "html" ? "（HTML格式）" : ""}`;
    case "set_paragraph_style":
      return `将段落${params.paragraph_index}设置为${params.format}格式`;
    case "merge_paragraphs":
      return `合并段落${params.first_paragraph_index}和段落${params.second_paragraph_index}${params.separator ? `（分隔符: "${params.separator}"）` : ""}`;
    case "reply_only":
      return "仅回复文本";
    case "task_complete":
      return `任务完成：${params.summary || ""}`;
    case "undo_last_action":
      return `撤销最近 ${params.count || 1} 步操作`;
    case "insert_mermaid_image":
      return `在文档中插入 mermaid 流程图（位置: ${params.location ?? "at_cursor"}）`;
    default:
      return `${actionName}: ${JSON.stringify(params)}`;
  }
}

// --- Parse action plan from text (fallback for non-FC models) ---

export function parseActionPlanFromText(text: string): ActionPlan | null {
  // Try to extract <action_plan>...</action_plan> block
  const planMatch = text.match(/<action_plan>\s*([\s\S]*?)\s*<\/action_plan>/);
  if (!planMatch) {
    return null;
  }

  try {
    const plan = JSON.parse(planMatch[1].trim());
    if (!plan.actions || !Array.isArray(plan.actions)) {
      return null;
    }

    const timestamp = Date.now();
    const actions: WordAction[] = plan.actions.map((a: any, index: number) => ({
      action: a.action || a.name || "reply_only",
      params: a.params || a.parameters || {},
      description: a.description || describeAction(a.action || a.name || "reply_only", a.params || a.parameters || {}),
      toolCallId: `text-parsed-${timestamp}-${index}`,
    }));

    return {
      actions,
      explanation: plan.explanation || actions.map((a) => a.description).join("；"),
    };
  } catch {
    return null;
  }
}

// --- Extract pure text reply (strip <action_plan> blocks) ---

export function extractTextReply(text: string): string {
  return text.replace(/<action_plan>[\s\S]*?<\/action_plan>/g, "").trim();
}

// --- API helpers ---

function getEndpoint(config: ProviderConfig): string {
  return config.baseUrl.replace(/\/$/, "") + "/chat/completions";
}

function getModelsEndpoint(config: ProviderConfig): string {
  return config.baseUrl.replace(/\/$/, "") + "/models";
}

export async function listOpenAICompatibleModels(config: ProviderConfig): Promise<string[]> {
  const endpoint = getModelsEndpoint(config);

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${config.apiKey}`,
      },
      signal: AbortSignal.timeout(20_000),
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown network error";
    throw new Error(`LLM network error to ${endpoint}: ${message}`);
  }

  if (!response.ok) {
    const details = await response.text();
    throw new LlmHttpError(response.status, details);
  }

  const data = (await response.json()) as ModelsResponse;
  return (data.data ?? []).map((item) => item.id).filter((id): id is string => Boolean(id));
}

export async function callOpenAICompatible(
  config: ProviderConfig,
  messages: ChatMessage[],
  contextChunks: RetrievalChunk[],
  tools?: ToolDefinition[],
  insertMode?: string,
  documentStructureDescription?: string,
  dynamicContext?: { documentContext?: string; selection?: string }
): Promise<{ reply: string; actionPlan: ActionPlan | null; reasoningContent?: string; finishReason?: string; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> }> {
  const payload = buildPayload(config, messages, contextChunks, false, tools, insertMode, documentStructureDescription, dynamicContext);
  const endpoint = getEndpoint(config);
  const timeoutMs = Number(process.env.LLM_TIMEOUT_MS ?? 180_000);

  const traceId = await logRequestStart({
    endpoint: "callOpenAICompatible",
    model: config.model,
    messages: payload.messages,
    tools: payload.tools,
    temperature: payload.temperature,
    maxTokens: payload.max_tokens,
    stream: false,
    dynamicContext,
    retrievedChunks: contextChunks.map((c) => ({ id: c.id, fileName: c.fileName, text: c.text })),
  });

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "text/event-stream, application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(timeoutMs),
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown network error";
    const causeCode =
      error && typeof error === "object" && "cause" in error
        ? (error as { cause?: { code?: string } }).cause?.code
        : undefined;
    const codeSuffix = causeCode ? ` (${causeCode})` : "";
    await logError(traceId, { endpoint: "callOpenAICompatible", error: `LLM network error: ${message}${codeSuffix}` });
    throw new Error(`LLM network error to ${endpoint}: ${message}${codeSuffix}`);
  }

  if (!response.ok) {
    const details = await response.text();
    await logResponseRaw(traceId, { status: response.status, contentType: response.headers.get("content-type") ?? undefined, bodyPreview: details.slice(0, 2000) });
    await logError(traceId, { endpoint: "callOpenAICompatible", error: `HTTP ${response.status}: ${details.slice(0, 500)}` });
    throw new LlmHttpError(response.status, details);
  }

  const contentType = response.headers.get("content-type") || "";
  const rawText = await response.text();
  await logResponseRaw(traceId, { status: response.status, contentType, bodyPreview: rawText.slice(0, 4000) });

  let data: ChatCompletionResponse;
  try {
    data = JSON.parse(rawText) as ChatCompletionResponse;
  } catch (parseErr) {
    await logError(traceId, { endpoint: "callOpenAICompatible", error: `JSON parse failed: ${(parseErr as Error).message}. Raw length=${rawText.length}. First 200 chars: ${rawText.slice(0, 200)}` });
    throw new Error(`LLM response JSON parse failed: ${(parseErr as Error).message}`);
  }

  const choice = data.choices?.[0];
  const message = choice?.message;

  if (!message) {
    await logResponseParsed(traceId, { reply: "模型没有返回可用内容。", actionPlan: null });
    return { reply: "模型没有返回可用内容。", actionPlan: null };
  }

  const reasoningContent = message.reasoning_content || undefined;
  const finishReason = choice?.finish_reason;

  // Check for tool_calls
  if (message.tool_calls && message.tool_calls.length > 0) {
    const actionPlan = parseActionPlanFromToolCalls(message.tool_calls);
    // If there's also text content, include it as the reply
    const reply = message.content?.trim() || actionPlan.explanation;
    await logResponseParsed(traceId, { reply, actionPlan, toolCalls: message.tool_calls });
    return { reply, actionPlan, reasoningContent, finishReason, toolCalls: message.tool_calls };
  }

  // No tool_calls — check for text-based action plan (fallback)
  const textContent = message.content?.trim() || "模型没有返回可用内容。";
  const textActionPlan = parseActionPlanFromText(textContent);
  const reply = textActionPlan ? extractTextReply(textContent) || textActionPlan.explanation : textContent;

  await logResponseParsed(traceId, { reply, actionPlan: textActionPlan });
  return { reply, actionPlan: textActionPlan, reasoningContent, finishReason };
}

// Helper to check if an action plan contains only perception tools
export function isPerceptionOnlyPlan(plan: ActionPlan): boolean {
  const perceptionTools = ["read_document", "get_selection_info", "get_document_stats", "get_paragraph_format", "get_document_tables", "read_table"];
  return plan.actions.length > 0 && plan.actions.every((a) => perceptionTools.includes(a.action));
}

export function isIterablePlan(plan: ActionPlan): boolean {
  const perceptionTools = ["read_document", "get_selection_info", "get_document_stats", "get_paragraph_format", "get_document_tables", "read_table"];
  const iterableActionTools = [
    "find_and_replace", "find_and_replace_v2",
    "insert_at_cursor", "insert_after_heading", "insert_at_end", "insert_after_paragraph",
    "insert_table", "delete_table",
    "insert_mermaid_image",
    "delete_paragraph",
    "replace_paragraph", "set_paragraph_style", "merge_paragraphs", "apply_rich_format",
    "undo_last_action",
  ];
  const iterableTools = [...perceptionTools, ...iterableActionTools];
  return plan.actions.length > 0 && plan.actions.every((a) => iterableTools.includes(a.action));
}

export async function streamOpenAICompatible(
  config: ProviderConfig,
  messages: ChatMessage[],
  contextChunks: RetrievalChunk[],
  onDelta: (text: string) => void,
  signal?: AbortSignal,
  tools?: ToolDefinition[],
  insertMode?: string,
  documentStructureDescription?: string,
  dynamicContext?: { documentContext?: string; selection?: string }
): Promise<{ reply: string; actionPlan: ActionPlan | null; reasoningContent?: string; finishReason?: string; toolCalls?: Array<{ id: string; type: "function"; function: { name: string; arguments: string } }> }> {
  const payload = buildPayload(config, messages, contextChunks, true, tools, insertMode, documentStructureDescription, dynamicContext);
  const endpoint = getEndpoint(config);
  const overallTimeoutMs = config.overallTimeout ? config.overallTimeout * 1000 : Number(process.env.LLM_STREAM_TIMEOUT_MS ?? 240_000);
  const firstTokenTimeoutMs = config.firstTokenTimeout ? config.firstTokenTimeout * 1000 : Number(process.env.LLM_STREAM_FIRST_TOKEN_TIMEOUT_MS ?? 20_000);

  const traceId = await logRequestStart({
    endpoint: "streamOpenAICompatible",
    model: config.model,
    messages: payload.messages,
    tools: payload.tools,
    temperature: payload.temperature,
    maxTokens: payload.max_tokens,
    stream: true,
    dynamicContext,
    retrievedChunks: contextChunks.map((c) => ({ id: c.id, fileName: c.fileName, text: c.text })),
  });

  const controller = new AbortController();
  let externalAbortHandler: (() => void) | null = null;
  if (signal) {
    if (signal.aborted) {
      controller.abort(signal.reason);
    } else {
      externalAbortHandler = () => controller.abort(signal.reason);
      signal.addEventListener("abort", externalAbortHandler, { once: true });
    }
  }

  const overallTimer = setTimeout(() => {
    controller.abort(new Error("stream_overall_timeout"));
  }, overallTimeoutMs);

  let firstTokenReceived = false;
  let firstTokenTimer = setTimeout(() => {
    if (!firstTokenReceived) {
      controller.abort(new Error("stream_first_token_timeout"));
    }
  }, firstTokenTimeoutMs);

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });
  } catch (error) {
    clearTimeout(firstTokenTimer);
    clearTimeout(overallTimer);
    if (signal && externalAbortHandler) {
      signal.removeEventListener("abort", externalAbortHandler);
    }

    if (error instanceof Error && controller.signal.aborted) {
      const reason = controller.signal.reason;
      if (reason instanceof Error && reason.message === "stream_first_token_timeout") {
        await logError(traceId, { endpoint: "streamOpenAICompatible", error: `first token timeout ${firstTokenTimeoutMs}ms` });
        throw new Error(`LLM stream timeout: first token not received within ${firstTokenTimeoutMs}ms`);
      }
      if (reason instanceof Error && reason.message === "stream_overall_timeout") {
        await logError(traceId, { endpoint: "streamOpenAICompatible", error: `overall timeout ${overallTimeoutMs}ms` });
        throw new Error(`LLM stream timeout: response not finished within ${overallTimeoutMs}ms`);
      }
    }

    const message = error instanceof Error ? error.message : "Unknown network error";
    const causeCode =
      error && typeof error === "object" && "cause" in error
        ? (error as { cause?: { code?: string } }).cause?.code
        : undefined;
    const codeSuffix = causeCode ? ` (${causeCode})` : "";
    await logError(traceId, { endpoint: "streamOpenAICompatible", error: `network error: ${message}${codeSuffix}` });
    throw new Error(`LLM network error to ${endpoint}: ${message}${codeSuffix}`);
  }

  if (!response.ok) {
    clearTimeout(firstTokenTimer);
    clearTimeout(overallTimer);
    if (signal && externalAbortHandler) {
      signal.removeEventListener("abort", externalAbortHandler);
    }
    const details = await response.text();
    await logResponseRaw(traceId, { status: response.status, contentType: response.headers.get("content-type") ?? undefined, bodyPreview: details.slice(0, 2000) });
    await logError(traceId, { endpoint: "streamOpenAICompatible", error: `HTTP ${response.status}: ${details.slice(0, 500)}` });
    throw new LlmHttpError(response.status, details);
  }

  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    clearTimeout(firstTokenTimer);
    clearTimeout(overallTimer);
    if (signal && externalAbortHandler) {
      signal.removeEventListener("abort", externalAbortHandler);
    }

    const rawText = await response.text();
    await logResponseRaw(traceId, { status: response.status, contentType, bodyPreview: rawText.slice(0, 4000) });

    let data: ChatCompletionResponse;
    try {
      data = JSON.parse(rawText) as ChatCompletionResponse;
    } catch (parseErr) {
      await logError(traceId, { endpoint: "streamOpenAICompatible", error: `JSON parse failed (json branch): ${(parseErr as Error).message}. Raw length=${rawText.length}. First 200 chars: ${rawText.slice(0, 200)}` });
      throw new Error(`LLM response JSON parse failed: ${(parseErr as Error).message}`);
    }

    const choice = data.choices?.[0];
    const message = choice?.message;

    if (!message) {
      await logResponseParsed(traceId, { reply: "模型没有返回可用内容。", actionPlan: null });
      return { reply: "模型没有返回可用内容。", actionPlan: null };
    }

    const reasoningContent = message.reasoning_content || undefined;
    const finishReason = choice?.finish_reason;

    // Check for tool_calls in non-streaming JSON response
    if (message.tool_calls && message.tool_calls.length > 0) {
      const actionPlan = parseActionPlanFromToolCalls(message.tool_calls);
      const reply = message.content?.trim() || actionPlan.explanation;
      if (reply) {
        onDelta(reply);
      }
      await logResponseParsed(traceId, { reply, actionPlan, toolCalls: message.tool_calls });
      return { reply, actionPlan, reasoningContent, finishReason, toolCalls: message.tool_calls };
    }

    const textContent = message.content?.trim() || "模型没有返回可用内容。";
    const textActionPlan = parseActionPlanFromText(textContent);
    const reply = textActionPlan ? extractTextReply(textContent) || textActionPlan.explanation : textContent;
    if (textContent) {
      onDelta(textContent);
    }
    await logResponseParsed(traceId, { reply, actionPlan: textActionPlan });
    return { reply, actionPlan: textActionPlan, reasoningContent, finishReason };
  }

  if (!response.body) {
    clearTimeout(firstTokenTimer);
    clearTimeout(overallTimer);
    if (signal && externalAbortHandler) {
      signal.removeEventListener("abort", externalAbortHandler);
    }
    throw new Error("LLM stream response has no body.");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let buffer = "";
  let fullText = "";
  let fullReasoning = "";
  let lastFinishReason: string | undefined;
  let pendingToolCalls: Map<number, { id: string; name: string; arguments: string }> = new Map();

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split(/\r?\n/);
    buffer = lines.pop() ?? "";

    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line.startsWith("data:")) {
        continue;
      }

      const data = line.slice(5).trim();
      if (!data || data === "[DONE]") {
        continue;
      }

      let parsed: StreamChunk;
      try {
        parsed = JSON.parse(data) as StreamChunk;
      } catch {
        continue;
      }

      // Capture finish_reason from the choice — OpenAI sends it on the final chunk
      const choiceFinishReason = parsed.choices?.[0]?.finish_reason;
      if (choiceFinishReason) {
        lastFinishReason = choiceFinishReason;
      }

      const delta = parsed.choices?.[0]?.delta;
      if (!delta) {
        // Try legacy format
        const text = parsed.choices?.[0]?.text ?? parsed.choices?.[0]?.message?.content ?? "";
        if (text) {
          if (!firstTokenReceived) {
            firstTokenReceived = true;
            clearTimeout(firstTokenTimer);
          }
          fullText += text;
          onDelta(text);
        }
        continue;
      }

      // Helper to mark first token received (any non-empty delta counts)
      const markFirstToken = () => {
        if (!firstTokenReceived) {
          firstTokenReceived = true;
          clearTimeout(firstTokenTimer);
        }
      };

      // Handle tool_calls in streaming
      if (delta.tool_calls && Array.isArray(delta.tool_calls)) {
        markFirstToken();
        for (const tc of delta.tool_calls as Array<{ index?: number; id?: string; function?: { name?: string; arguments?: string } }>) {
          const idx = tc.index ?? 0;
          const existing = pendingToolCalls.get(idx);
          if (tc.id) {
            // New tool call starting
            pendingToolCalls.set(idx, {
              id: tc.id,
              name: tc.function?.name || (existing?.name ?? ""),
              arguments: tc.function?.arguments || "",
            });
          } else if (existing) {
            // Continuation of tool call
            if (tc.function?.name) {
              existing.name += tc.function.name;
            }
            if (tc.function?.arguments) {
              existing.arguments += tc.function.arguments;
            }
          }
        }
        continue;
      }

      // Handle reasoning content (thinking chain in streaming mode)
      const reasoning = delta.reasoning_content ?? "";
      if (reasoning) {
        markFirstToken();
        fullReasoning += reasoning;
      }

      // Handle text content
      const content = delta.content ?? "";
      if (content) {
        markFirstToken();
        fullText += content;
        onDelta(content);
      }
    }
  }

  clearTimeout(firstTokenTimer);
  clearTimeout(overallTimer);
  if (signal && externalAbortHandler) {
    signal.removeEventListener("abort", externalAbortHandler);
  }

  await logResponseRaw(traceId, { status: response.status, contentType, streamChunkCount: pendingToolCalls.size, fullTextLength: fullText.length, fullReasoningLength: fullReasoning.length });

  // Process accumulated tool calls
  if (pendingToolCalls.size > 0) {
    const toolCalls = [...pendingToolCalls.entries()]
      .sort(([a], [b]) => a - b)
      .map(([, tc]) => ({
        id: tc.id,
        type: "function" as const,
        function: {
          name: tc.name,
          arguments: tc.arguments,
        },
      }));

    const actionPlan = parseActionPlanFromToolCalls(toolCalls);
    const reply = fullText.trim() || actionPlan.explanation;
    await logResponseParsed(traceId, { reply, actionPlan, toolCalls });
    return { reply, actionPlan, reasoningContent: fullReasoning || undefined, finishReason: lastFinishReason, toolCalls };
  }

  // No tool calls — check for text-based action plan.
  // reasoning_content is internal monologue (DeepSeek, Agnes, etc.); never
  // used as the user-facing reply.  The loop decision is driven by
  // finish_reason, not by whether reasoning_content is present.
  const textActionPlan = parseActionPlanFromText(fullText);
  const reply = textActionPlan ? extractTextReply(fullText) || textActionPlan.explanation : fullText.trim() || "模型没有返回可用内容。";
  await logResponseParsed(traceId, { reply, actionPlan: textActionPlan });
  return { reply, actionPlan: textActionPlan, reasoningContent: fullReasoning || undefined, finishReason: lastFinishReason };
}
