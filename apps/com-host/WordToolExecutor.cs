using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeAgent.ComHost
{
    /// <summary>
    /// WP3 / D3: executes the document tools in Word OM on behalf of the TS UI
    /// (behavior transliterated from apps/word-addin/src/main.ts executeAction
    /// and the perception helpers). Word is late-bound via dynamic/IDispatch —
    /// no Word PIA required at build or runtime.
    /// </summary>
    internal static class WordToolExecutor
    {
        internal static dynamic App;

        internal sealed class ToolResult
        {
            public bool Success;
            public string Result;
            public object Data;
            public string Error;
        }

        private static readonly Dictionary<string, object[]> StyleMap = new Dictionary<string, object[]>
        {
            // Built-in style IDs first (locale-proof), then localized names.
            { "normal", new object[] { -1, "Normal", "正文" } },
            { "heading1", new object[] { -2, "Heading 1", "标题 1", "标题1" } },
            { "heading2", new object[] { -3, "Heading 2", "标题 2", "标题2" } },
            { "heading3", new object[] { -4, "Heading 3", "标题 3", "标题3" } },
            { "bullet_list", new object[] { -67, "List Paragraph", "列表段落" } },
            { "numbered_list", new object[] { -67, "List Paragraph", "列表段落" } },
        };

        private const int MaxStructuredParagraphs = 80;
        private const int MaxStructuredText = 200;

        // ── protocol entry points ─────────────────────────────────────────────

        internal static ToolResult Execute(string method, JObject payload)
        {
            try
            {
                switch (method)
                {
                    case "executeAction":
                        return RunAction((JObject)payload["action"]);
                    case "getStructuredContext":
                        return new ToolResult { Success = true, Data = BuildStructuredContext() };
                    case "getWordContext":
                        return new ToolResult { Success = true, Data = BuildWordContext() };
                    case "simpleInsert":
                        return SimpleInsert((string)payload["mode"], (string)payload["text"] ?? "");
                    default:
                        return new ToolResult { Success = false, Error = "未知桥方法: " + method };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Success = false, Error = ex.Message };
            }
        }

        // ── action dispatcher (mirrors main.ts executeAction) ─────────────────

        private static ToolResult RunAction(JObject action)
        {
            string name = (string)action["action"] ?? "";
            JObject p = (JObject)action["params"] ?? new JObject();

            try
            {
                string result = DispatchAction(name, p);
                return new ToolResult { Success = true, Result = result };
            }
            catch (Exception ex)
            {
                return new ToolResult { Success = false, Error = ex.Message };
            }
        }

        private static string DispatchAction(string name, JObject p)
        {
            dynamic doc = App.ActiveDocument;
            switch (name)
            {
                case "reply_only":
                    return "仅回复文本，无需文档操作";
                case "task_complete":
                    return (string)p["summary"] ?? "任务已完成";

                case "read_document":
                    return ReadDocument(p);
                case "get_selection_info":
                    return GetSelectionInfo();
                case "get_document_stats":
                    return GetDocumentStats();
                case "get_paragraph_format":
                    return GetParagraphFormat(p);
                case "get_document_tables":
                    return GetDocumentTables();
                case "read_table":
                    return ReadTable(p);

                case "insert_after_heading":
                    return InsertAfterHeading(doc, p);
                case "replace_selection":
                    return InsertIntoSelection(doc, p, "已替换选区内容");
                case "insert_at_cursor":
                    return InsertIntoSelection(doc, p, "已在光标处插入内容");
                case "insert_at_end":
                    return InsertAtEnd(doc, p);                case "insert_at_start":
                    return InsertAtStart(doc, p);
                case "insert_after_paragraph":
                    return InsertAfterParagraph(doc, p);
                case "delete_paragraph":
                    return DeleteParagraph(doc, p);
                case "replace_paragraph":
                    return ReplaceParagraph(doc, p);
                case "set_paragraph_style":
                    return SetParagraphStyle(doc, p);
                case "merge_paragraphs":
                    return MergeParagraphs(doc, p);
                case "apply_rich_format":
                    return ApplyRichFormat(doc, p);
                case "find_and_replace":
                    return FindAndReplace(doc, p, false);
                case "find_and_replace_v2":
                    return FindAndReplace(doc, p, true);
                case "undo_last_action":
                    return UndoLastAction(doc, p);
                case "delete_table":
                    return DeleteTable(doc, p);
                case "insert_table":
                    return InsertTable(doc, p);
                case "insert_mermaid_image":
                    return InsertMermaidImage(doc, p);
                default:
                    return "未知操作: " + name;
            }
        }

        // ── perception tools ──────────────────────────────────────────────────

        private static string ReadDocument(JObject p)
        {
            dynamic doc = App.ActiveDocument;
            string mode = (string)p["mode"];
            int total = (int)doc.Paragraphs.Count;

            if (mode == "paragraph_range")
            {
                int start = p["paragraph_index"] != null ? (int)p["paragraph_index"] : 0;
                int count = p["count"] != null ? (int)p["count"] : 5;
                var lines = new List<string>();
                for (int i = start; i < Math.Min(start + count, total); i++)
                {
                    lines.Add(DescribeParagraph(ParagraphAt(doc, i), i));
                }
                return lines.Count > 0
                    ? string.Format("读取段落 {0} 起共 {1} 段（总计 {2} 段）：\n{3}", start, lines.Count, total, string.Join("\n", lines))
                    : "未读取到段落内容";
            }

            if (mode == "heading_context")
            {
                string headingText = (string)p["heading_text"];
                if (string.IsNullOrEmpty(headingText)) throw new Error("heading_context 模式需要提供 heading_text");
                return HeadingContext(doc, headingText);
            }

            if (mode == "selection")
            {
                string text = SafeText(App.Selection.Text);
                if (text.Length > 2000) text = text.Substring(0, 2000);
                return text.Length > 0 ? "当前选区内容：" + text : "当前无选中文本";
            }

            if (mode == "cursor_surrounding")
            {
                int chars = p["surrounding_chars"] != null ? (int)p["surrounding_chars"] : 500;
                string selText = SafeText(App.Selection.Text);
                int mid = selText.Length / 2;
                int start = Math.Max(0, mid - chars);
                int end = Math.Min(selText.Length, mid + chars);
                string surrounding = selText.Substring(start, Math.Max(0, end - start));
                return "光标周围内容：" + (string.IsNullOrEmpty(surrounding) ? "（无法获取光标周围内容）" : surrounding);
            }

            return "未知的读取模式: " + mode;
        }

        private static string DescribeParagraph(dynamic para, int index)
        {
            string text = SafeText(para.Range.Text).TrimEnd('\r', '\a');
            int imgCount = para.Range.InlineShapes.Count;
            string styleName = StyleNameOf(para);
            int headingLevel = HeadingLevelOf(styleName);
            var fontParts = new List<string>();
            dynamic font = para.Range.Font;
            string fontName = font.Name as string;
            if (!string.IsNullOrEmpty(fontName)) fontParts.Add(fontName);
            double fontSize = (double)font.Size;
            if (fontSize > 0) fontParts.Add(fontSize.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "pt");
            if (font.Bold != 0) fontParts.Add("加粗");
            if (font.Italic != 0) fontParts.Add("斜体");
            string colorHex = ColorHex((int)font.Color);
            if (colorHex != null && colorHex.ToUpperInvariant() != "#000000") fontParts.Add(colorHex);
            string fontStr = fontParts.Count > 0 ? " [" + string.Join(" ", fontParts) + "]" : "";
            string imgStr = imgCount > 0 ? " [📷图片×" + imgCount + "]" : "";
            string displayText = text.Length > 0 ? text : (imgCount > 0 ? "（图片段落，无文字）" : "");
            return string.Format("[段落{0}] {1}{2}{3}{4}", index,
                headingLevel > 0 ? "标题" + headingLevel + " " : "", displayText, imgStr, fontStr);
        }

        private static string HeadingContext(dynamic doc, string headingText)
        {
            int total = (int)doc.Paragraphs.Count;
            var matchInfos = new List<int[]>(); // [paraIndex, isHeading(0/1), level]
            var matchTexts = new List<string>();
            for (int i = 0; i < total; i++)
            {
                dynamic para = ParagraphAt(doc, i);
                string text = SafeText(para.Range.Text).TrimEnd('\r', '\a').Trim();
                if (text.IndexOf(headingText, StringComparison.OrdinalIgnoreCase) < 0) continue;
                int level = HeadingLevelOf(StyleNameOf(para));
                matchInfos.Add(new[] { i, level > 0 ? 1 : 0, level });
                matchTexts.Add(text);
            }
            if (matchInfos.Count == 0) return "未找到标题\"" + headingText + "\"";

            var heading = matchInfos.FirstOrDefault(m => m[1] == 1) ?? matchInfos[0];
            string multiMatchNote = "";
            if (matchInfos.Count > 1)
            {
                var matchList = new List<string>();
                for (int k = 0; k < matchInfos.Count; k++)
                {
                    matchList.Add(string.Format("  [段落{0}] {1}: {2}", matchInfos[k][0],
                        matchInfos[k][1] == 1 ? "标题" + matchInfos[k][2] : "普通文本", matchTexts[k]));
                }
                multiMatchNote = string.Format("\n⚠️ 找到 {0} 处匹配，已自动选择{1}（段落{2}）。所有匹配：\n{3}\n如需查看其他匹配项，请使用 read_document 的 paragraph_range 模式指定具体段落序号。\n",
                    matchInfos.Count, matchInfos.Any(m => m[1] == 1) ? "实际标题" : "第一个匹配", heading[0], string.Join("\n", matchList));
            }

            int startLevel = heading[2];
            bool found = false;
            var result = new List<string>();
            for (int i = 0; i < total; i++)
            {
                dynamic para = ParagraphAt(doc, i);
                string text = SafeText(para.Range.Text).TrimEnd('\r', '\a');
                int level = HeadingLevelOf(StyleNameOf(para));
                int imgCount = para.Range.InlineShapes.Count;
                string imgStr = imgCount > 0 ? " [📷图片×" + imgCount + "]" : "";
                if (!found)
                {
                    if (i == heading[0])
                    {
                        found = true;
                        result.Add(string.Format("[段落{0}] 标题{1} {2}{3}", i, level, text, imgStr));
                    }
                    continue;
                }
                if (level > 0 && level <= startLevel) break;
                string displayText = text.Length > 0 ? text : (imgCount > 0 ? "（图片段落，无文字）" : "");
                result.Add(string.Format("[段落{0}] {1}{2}", i, displayText, imgStr));
            }
            return string.Format("标题\"{0}\"及其子内容：{1}\n{2}", headingText, multiMatchNote, string.Join("\n", result));
        }

        private static string GetSelectionInfo()
        {
            dynamic doc = App.ActiveDocument;
            dynamic sel = App.Selection;
            string text = SafeText(sel.Text);
            bool isCursorOnly = text.Length == 0;
            int paragraphCount = sel.Paragraphs.Count;

            int? startIdx = null, endIdx = null;
            if (paragraphCount > 0)
            {
                string firstSel = SafeText(sel.Paragraphs[1].Range.Text).TrimEnd('\r', '\a').Trim();
                string lastSel = SafeText(sel.Paragraphs[paragraphCount].Range.Text).TrimEnd('\r', '\a').Trim();
                int total = (int)doc.Paragraphs.Count;
                for (int i = 0; i < total; i++)
                {
                    string pText = SafeText(ParagraphAt(doc, i).Range.Text).TrimEnd('\r', '\a').Trim();
                    if (startIdx == null && (pText == firstSel || (firstSel.Length > 0 && pText.StartsWith(firstSel)))) startIdx = i;
                    if (endIdx == null && (pText == lastSel || (lastSel.Length > 0 && pText.StartsWith(lastSel)))) endIdx = i;
                }
            }

            var info = new JObject
            {
                ["isCursorOnly"] = isCursorOnly,
                ["textLength"] = text.Length,
                ["text"] = text.Length > 200 ? text.Substring(0, 200) : text,
                ["paragraphCount"] = paragraphCount,
            };
            if (startIdx != null) info["startParagraphIndex"] = startIdx.Value;
            if (endIdx != null) info["endParagraphIndex"] = endIdx.Value;
            return info.ToString(Formatting.None);
        }

        private static string GetDocumentStats()
        {
            dynamic doc = App.ActiveDocument;
            int total = (int)doc.Paragraphs.Count;
            var headings = new JArray();
            int listCount = 0;
            for (int i = 0; i < total; i++)
            {
                dynamic para = ParagraphAt(doc, i);
                int level = HeadingLevelOf(StyleNameOf(para));
                if (level > 0)
                {
                    string text = SafeText(para.Range.Text).TrimEnd('\r', '\a').Trim();
                    if (text.Length > 100) text = text.Substring(0, 100);
                    headings.Add(new JObject { ["level"] = level, ["text"] = text, ["paragraphIndex"] = i });
                }
                if (para.Range.ListFormat.ListType != 0) listCount++;
            }
            var stats = new JObject
            {
                ["totalParagraphs"] = total,
                ["totalCharacters"] = ((string)doc.Content.Text ?? "").Length,
                ["headingCount"] = headings.Count,
                ["listParagraphCount"] = listCount,
                ["headings"] = headings,
            };
            return stats.ToString(Formatting.None);
        }

        private static string GetParagraphFormat(JObject p)
        {
            dynamic doc = App.ActiveDocument;
            int index = (int)p["paragraph_index"];
            int total = (int)doc.Paragraphs.Count;
            if (index < 0 || index >= total) throw new Error(string.Format("段落序号 {0} 超出范围（共 {1} 段）", index, total));

            dynamic para = ParagraphAt(doc, index);
            string text = SafeText(para.Range.Text).TrimEnd('\r', '\a');
            string styleName = StyleNameOf(para);
            int headingLevel = HeadingLevelOf(styleName);
            bool isListItem = para.Range.ListFormat.ListType != 0;

            var paragraphFont = new JObject();
            dynamic font = para.Range.Font;
            string fontName = font.Name as string;
            if (!string.IsNullOrEmpty(fontName)) paragraphFont["name"] = fontName;
            double fontSize = (double)font.Size;
            if (fontSize > 0) paragraphFont["size"] = fontSize;
            string colorHex = ColorHex((int)font.Color);
            if (colorHex != null) paragraphFont["color"] = colorHex;
            if (font.Bold != 0) paragraphFont["bold"] = true;
            if (font.Italic != 0) paragraphFont["italic"] = true;

            JArray runs = BuildFormatRuns(doc, para, text);

            var signatures = new HashSet<string>();
            foreach (JObject runInfo in runs)
            {
                signatures.Add(((JObject)runInfo["font"] ?? new JObject()).ToString(Formatting.None));
            }

            var result = new JObject
            {
                ["paragraphIndex"] = index,
                ["text"] = text.Length > 1000 ? text.Substring(0, 1000) : text,
                ["style"] = styleName,
                ["isListItem"] = isListItem,
                ["hasMixedFormatting"] = signatures.Count > 1,
                ["runs"] = runs,
            };
            if (headingLevel > 0) result["headingLevel"] = headingLevel;
            if (paragraphFont.Count > 0) result["paragraphFont"] = paragraphFont;
            return result.ToString(Formatting.None);
        }

        /// <summary>
        /// Splits paragraph text into formatting runs. Word's Range.Runs is a
        /// parameterized COM property that neither the C# dynamic binder nor
        /// InvokeMember can read, so run boundaries are detected by comparing
        /// per-character font signatures (equivalent to the span parsing the
        /// TS implementation did on Word's HTML output).
        /// </summary>
        private static JArray BuildFormatRuns(dynamic doc, dynamic para, string text)
        {
            var runs = new JArray();
            if (text.Length == 0) return runs;

            int absStart = (int)para.Range.Start;
            if (text.Length > 3000)
            {
                runs.Add(new JObject { ["text"] = text.Substring(0, 200) + "..." });
                return runs;
            }

            int runStart = 0;
            string sig = null;
            for (int i = 0; i < text.Length; i++)
            {
                dynamic cr = doc.Range(absStart + i, absStart + i + 1);
                string s = FontSignature(cr);
                if (sig == null) { sig = s; continue; }
                if (s != sig)
                {
                    AddRun(runs, text.Substring(runStart, i - runStart), sig);
                    runStart = i;
                    sig = s;
                }
            }
            AddRun(runs, text.Substring(runStart), sig);
            return runs;
        }

        private static void AddRun(JArray runs, string runText, string sig)
        {
            if (runText.Trim().Length == 0) return;
            var runInfo = new JObject { ["text"] = runText.Length > 200 ? runText.Substring(0, 200) + "..." : runText };
            string[] parts = sig.Split('|');
            var runFont = new JObject();
            if (parts[0].Length > 0) runFont["name"] = parts[0];
            if (parts[1].Length > 0) runFont["size"] = parts[1];
            if (parts[2] == "1") runFont["bold"] = true;
            if (parts[3] == "1") runFont["italic"] = true;
            if (parts[4].Length > 0) runFont["color"] = parts[4];
            if (runFont.Count > 0) runInfo["font"] = runFont;
            runs.Add(runInfo);
        }

        private static string FontSignature(dynamic range)
        {
            dynamic font = range.Font;
            string name = (font.Name as string) ?? "";
            string size = ((double)font.Size).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            string bold = font.Bold != 0 ? "1" : "0";
            string italic = font.Italic != 0 ? "1" : "0";
            string color = ColorHex((int)font.Color) ?? "";
            return name + "|" + size + "|" + bold + "|" + italic + "|" + color;
        }

        private static string GetDocumentTables()
        {
            dynamic doc = App.ActiveDocument;
            int tableCount = doc.Tables.Count;
            if (tableCount == 0) return "文档中没有 Word 表格。";

            var paragraphTexts = new List<string>();
            int total = (int)doc.Paragraphs.Count;
            for (int i = 0; i < total; i++)
            {
                paragraphTexts.Add(SafeText(ParagraphAt(doc, i).Range.Text).TrimEnd('\r', '\a').Trim());
            }

            var tableInfos = new JArray();
            for (int t = 1; t <= tableCount; t++)
            {
                dynamic table = doc.Tables[t];
                int rowCount = table.Rows.Count;
                int colCount = table.Columns.Count;
                string style = "";
                try { style = StyleNameOf(table); } catch { }

                int? firstIdx = null, lastIdx = null;
                string preview = null;
                try
                {
                    string firstText = SafeText(table.Cell(1, 1).Range.Text).TrimEnd('\r', '\a').Trim();
                    preview = firstText.Length > 60 ? firstText.Substring(0, 60) : firstText;
                    firstIdx = paragraphTexts.IndexOf(firstText);
                    string lastText = SafeText(table.Cell(rowCount, colCount).Range.Text).TrimEnd('\r', '\a').Trim();
                    lastIdx = paragraphTexts.LastIndexOf(lastText);
                }
                catch { }

                var info = new JObject
                {
                    ["tableIndex"] = t - 1,
                    ["rowCount"] = rowCount,
                    ["columnCount"] = colCount,
                    ["style"] = style,
                };
                if (firstIdx != null && firstIdx.Value >= 0) info["firstParagraphIndex"] = firstIdx.Value;
                if (lastIdx != null && lastIdx.Value >= 0) info["lastParagraphIndex"] = lastIdx.Value;
                if (!string.IsNullOrEmpty(preview)) info["firstCellPreview"] = preview;
                tableInfos.Add(info);
            }

            var summary = new List<string>();
            foreach (JObject t in tableInfos)
            {
                string span = t["firstParagraphIndex"] != null && t["lastParagraphIndex"] != null
                    ? string.Format("（位于段落 {0} ~ {1}）", (int)t["firstParagraphIndex"], (int)t["lastParagraphIndex"])
                    : "";
                string prev = t["firstCellPreview"] != null ? "，首单元格：\"" + (string)t["firstCellPreview"] + "\"" : "";
                summary.Add(string.Format("  表格 {0}：{1} 行 × {2} 列，样式={3}{4}{5}",
                    (int)t["tableIndex"], (int)t["rowCount"], (int)t["columnCount"],
                    string.IsNullOrEmpty((string)t["style"]) ? "(默认)" : (string)t["style"], span, prev));
            }
            return string.Format("文档中共有 {0} 张表格：\n{1}\n\n--- 结构化数据（便于精确引用） ---\n{2}",
                tableCount, string.Join("\n", summary), tableInfos.ToString(Formatting.Indented));
        }

        private static string ReadTable(JObject p)
        {
            dynamic doc = App.ActiveDocument;
            int tableIndex = p["table_index"] != null ? (int)p["table_index"] : 0;
            int tableCount = doc.Tables.Count;
            if (tableIndex < 0 || tableIndex >= tableCount)
            {
                throw new Error(string.Format("表格序号 {0} 越界（文档共有 {1} 张表格，序号范围 0 ~ {2}）。请先调用 get_document_tables 确认有效索引。",
                    tableIndex, tableCount, tableCount - 1));
            }

            dynamic table = doc.Tables[tableIndex + 1];
            int rowCount = table.Rows.Count;
            int colCount = table.Columns.Count;
            string style = "";
            try { style = StyleNameOf(table); } catch { }

            var values = new JArray();
            for (int r = 1; r <= rowCount; r++)
            {
                var row = new JArray();
                for (int c = 1; c <= colCount; c++)
                {
                    string cellText = "";
                    try { cellText = SafeText(table.Cell(r, c).Range.Text).TrimEnd('\r', '\a'); } catch { }
                    row.Add(cellText);
                }
                values.Add(row);
            }

            var lines = new List<string>();
            lines.Add(string.Format("表格 {0}（{1} 行 × {2} 列，样式: {3}）：", tableIndex, rowCount, colCount,
                string.IsNullOrEmpty(style) ? "(默认)" : style));
            for (int r = 0; r < values.Count; r++)
            {
                var cells = new List<string>();
                var row = (JArray)values[r];
                for (int c = 0; c < row.Count; c++)
                {
                    cells.Add(string.Format("[列{0}] {1}", c + 1, (string)row[c] ?? ""));
                }
                lines.Add(string.Format("[行{0}] {1}", r + 1, string.Join("  ", cells)));
            }
            lines.Add("");
            lines.Add("--- JSON 格式（便于精确解析）---");
            lines.Add(new JObject
            {
                ["tableIndex"] = tableIndex,
                ["rowCount"] = rowCount,
                ["columnCount"] = colCount,
                ["style"] = style,
                ["values"] = values,
            }.ToString(Formatting.Indented));
            return string.Join("\n", lines);
        }

        // ── write tools ───────────────────────────────────────────────────────

        private static string InsertAfterHeading(dynamic doc, JObject p)
        {
            string headingText = (string)p["heading_text"];
            string content = (string)p["content"] ?? "";
            string format = (string)p["format"] ?? "normal";
            bool isHtml = (string)p["content_format"] == "html";
            if (string.IsNullOrEmpty(content.Trim())) throw new Error("插入内容为空或仅包含空白字符");

            int paraIndex = FindHeadingParagraphIndex(doc, headingText);
            if (paraIndex >= 0)
            {
                int pos = ParagraphAt(doc, paraIndex).Range.End;
                if (isHtml) { InsertHtmlAt(doc, pos, content); return "已在\"" + headingText + "\"后插入 HTML 格式内容"; }
                int lines = InsertMultiParagraphAt(doc, pos, content, format);
                return string.Format("已在\"{0}\"后插入内容（{1} 段）", headingText, lines);
            }

            int endPos = (int)doc.Content.End - 1;
            if (isHtml) { InsertHtmlAt(doc, endPos, content); return "未找到标题\"" + headingText + "\"，已插入 HTML 格式内容到文档末尾"; }
            int lineCount = InsertMultiParagraphAt(doc, endPos, content, format);
            return string.Format("未找到标题\"{0}\"，已插入到文档末尾（{1} 段）", headingText, lineCount);
        }

        private static string InsertAfterParagraph(dynamic doc, JObject p)
        {
            int index = (int)p["paragraph_index"];
            string content = (string)p["content"] ?? "";
            string format = (string)p["format"] ?? "normal";
            bool isHtml = (string)p["content_format"] == "html";
            if (string.IsNullOrEmpty(content.Trim())) throw new Error("插入内容为空或仅包含空白字符");
            int total = (int)doc.Paragraphs.Count;

            if (index >= 0 && index < total)
            {
                int pos = ParagraphAt(doc, index).Range.End;
                if (isHtml) { InsertHtmlAt(doc, pos, content); return string.Format("已在段落 {0} 后插入 HTML 格式内容", index); }
                int lines = InsertMultiParagraphAt(doc, pos, content, format);
                return string.Format("已在段落 {0} 后插入内容（{1} 段）", index, lines);
            }

            int endPos = (int)doc.Content.End - 1;
            if (isHtml) { InsertHtmlAt(doc, endPos, content); return string.Format("段落序号 {0} 超出范围，已插入 HTML 格式内容到末尾", index); }
            InsertMultiParagraphAt(doc, endPos, content, format);
            throw new Error(string.Format("insert_after_paragraph 失败：段落序号 {0} 超出范围（文档共 {1} 段）。请先调用 read_document 确认段落序号。", index, total));
        }

        private static string InsertAtEnd(dynamic doc, JObject p)
        {
            string content = (string)p["content"] ?? "";
            string format = (string)p["format"] ?? "normal";
            bool isHtml = (string)p["content_format"] == "html";
            if (string.IsNullOrEmpty(content.Trim())) throw new Error("插入内容为空或仅包含空白字符");
            int pos = (int)doc.Content.End - 1;
            if (isHtml) { InsertHtmlAt(doc, pos, content); return "已追加 HTML 格式内容到文档末尾"; }
            int lines = InsertMultiParagraphAt(doc, pos, content, format);
            return string.Format("已追加到文档末尾（{0} 段）", lines);
        }

        private static string InsertAtStart(dynamic doc, JObject p)
        {
            string content = (string)p["content"] ?? "";
            string format = (string)p["format"] ?? "normal";
            bool isHtml = (string)p["content_format"] == "html";
            if (string.IsNullOrEmpty(content.Trim())) throw new Error("插入内容为空或仅包含空白字符");
            int pos = doc.Paragraphs[1].Range.Start;
            if (isHtml) { InsertHtmlAt(doc, pos, content); return "已插入 HTML 格式内容到文档开头"; }
            int lines = InsertMultiParagraphAt(doc, pos, content, format);
            return string.Format("已插入到文档开头（{0} 段）", lines);
        }

        private static string InsertIntoSelection(dynamic doc, JObject p, string baseMessage)
        {
            string content = (string)p["content"] ?? "";
            string format = (string)p["format"] ?? "normal";
            bool isHtml = (string)p["content_format"] == "html";
            if (string.IsNullOrEmpty(content.Trim())) throw new Error("插入内容为空或仅包含空白字符");

            dynamic sel = App.Selection;
            if (isHtml)
            {
                dynamic r = sel.Range.Duplicate;
                r.Text = "";
                InsertHtmlAtRange(r, content);
                return baseMessage + "（HTML 格式）";
            }
            string[] lines = NormalizeContentLines(content);
            sel.Range.Text = string.Join("\r", lines);
            dynamic inserted = sel.Range;
            ApplyFormatToRange(inserted, format);
            return baseMessage == "已替换选区内容" ? baseMessage : string.Format("{0}（{1} 段）", baseMessage, lines.Length);
        }

        private static string DeleteParagraph(dynamic doc, JObject p)
        {
            int index = (int)p["paragraph_index"];
            int total = (int)doc.Paragraphs.Count;
            if (index < 0 || index >= total)
            {
                throw new Error(string.Format("delete_paragraph 失败：段落序号 {0} 超出范围（文档共 {1} 段）。请先调用 read_document 确认段落序号。", index, total));
            }
            try
            {
                ParagraphAt(doc, index).Range.Delete();
            }
            catch (Exception ex)
            {
                throw new Error(string.Format("delete_paragraph 失败：无法删除段落 {0}。可能原因：文档处于只读模式、受保护视图、或当前没有编辑权限。原始错误：{1}", index, ex.Message));
            }
            return "已删除段落 " + index;
        }

        private static string ReplaceParagraph(dynamic doc, JObject p)
        {
            int index = (int)p["paragraph_index"];
            string content = (string)p["content"] ?? "";
            bool isHtml = (string)p["content_format"] == "html";
            if (string.IsNullOrEmpty(content.Trim())) throw new Error("替换内容为空或仅包含空白字符");
            int total = (int)doc.Paragraphs.Count;
            if (index < 0 || index >= total) throw new Error(string.Format("段落序号 {0} 超出范围", index));

            dynamic para = ParagraphAt(doc, index);
            if (isHtml)
            {
                dynamic r = para.Range.Duplicate;
                r.Text = "";
                InsertHtmlAtRange(r, content);
                return string.Format("已替换段落 {0} 的内容（HTML 格式）", index);
            }
            string[] lines = NormalizeContentLines(content);
            SetParagraphText(para, lines[0]);
            if (lines.Length > 1)
            {
                int pos = para.Range.End;
                InsertMultiParagraphAt(doc, pos, string.Join("\r", lines.Skip(1)), "normal");
            }
            return string.Format("已替换段落 {0} 的内容（{1} 段）", index, lines.Length);
        }

        private static string SetParagraphStyle(dynamic doc, JObject p)
        {
            int index = (int)p["paragraph_index"];
            string format = (string)p["format"];
            int total = (int)doc.Paragraphs.Count;
            if (index < 0 || index >= total) throw new Error(string.Format("段落序号 {0} 超出范围", index));
            ApplyFormatToParagraph(ParagraphAt(doc, index), format);
            return string.Format("已将段落 {0} 设置为 {1} 格式", index, format);
        }

        private static string MergeParagraphs(dynamic doc, JObject p)
        {
            int first = (int)p["first_paragraph_index"];
            int second = (int)p["second_paragraph_index"];
            string separator = p["separator"] != null ? (string)p["separator"] : " ";
            int total = (int)doc.Paragraphs.Count;
            if (first < 0 || first >= total) throw new Error(string.Format("第一个段落序号 {0} 超出范围", first));
            if (second < 0 || second >= total) throw new Error(string.Format("第二个段落序号 {0} 超出范围", second));
            if (first == second) throw new Error("不能合并同一个段落");

            dynamic firstPara = ParagraphAt(doc, first);
            dynamic secondPara = ParagraphAt(doc, second);
            string merged = SafeText(firstPara.Range.Text).TrimEnd('\r', '\a') + separator + SafeText(secondPara.Range.Text).TrimEnd('\r', '\a');
            SetParagraphText(firstPara, merged);
            secondPara.Range.Delete();
            return string.Format("已合并段落 {0} 和段落 {1}", first, second);
        }

        private static string ApplyRichFormat(dynamic doc, JObject p)
        {
            string targetMode = (string)p["target_mode"];
            string textToFormat = (string)p["text_to_format"];
            dynamic targetRange;
            if (targetMode == "selection")
            {
                targetRange = App.Selection.Range.Duplicate;
            }
            else if (targetMode == "paragraph_index")
            {
                int index = (int)p["paragraph_index"];
                int total = (int)doc.Paragraphs.Count;
                if (index < 0 || index >= total) throw new Error(string.Format("段落序号 {0} 超出范围", index));
                targetRange = ParagraphAt(doc, index).Range.Duplicate;
            }
            else if (targetMode == "last_inserted")
            {
                throw new Error("last_inserted 模式暂未实现，请使用 paragraph_index 或 selection");
            }
            else
            {
                throw new Error("未知的目标模式: " + targetMode);
            }

            if (!string.IsNullOrEmpty(textToFormat))
            {
                dynamic search = targetRange.Duplicate;
                dynamic find = search.Find;
                find.ClearFormatting();
                bool hit = find.Execute(textToFormat, true, false); // matchCase=true
                if (!hit) throw new Error(string.Format("在目标范围内未找到文本\"{0}\"，请确认文本完全匹配（区分大小写）", textToFormat));
                targetRange = search;
            }

            JObject font = (JObject)p["font"];
            if (font != null)
            {
                dynamic f = targetRange.Font;
                if (font["name"] != null) f.Name = (string)font["name"];
                if (font["size"] != null) f.Size = (double)font["size"];
                if (font["color"] != null) f.Color = WdColorFromHex((string)font["color"]);
                if (font["bold"] != null) f.Bold = (bool)font["bold"] ? -1 : 0;
                if (font["italic"] != null) f.Italic = (bool)font["italic"] ? -1 : 0;
            }
            return textToFormat != null
                ? "已对文本\"" + textToFormat + "\"应用富文本格式"
                : "已应用富文本格式";
        }

        private static string FindAndReplace(dynamic doc, JObject p, bool v2)
        {
            string findText = (string)p["find_text"];
            string replaceText = (string)p["replace_text"] ?? "";
            bool matchCase = p["match_case"] != null && (bool)p["match_case"];
            bool matchWholeWord = v2 && p["match_whole_word"] != null && (bool)p["match_whole_word"];
            bool replaceAll = v2 && p["replace_all"] != null && (bool)p["replace_all"];

            // Count matches first (same reporting as the TS implementation).
            int total = 0;
            {
                dynamic rng = doc.Content.Duplicate;
                while (true)
                {
                    dynamic find = rng.Find;
                    find.ClearFormatting();
                    bool hit = find.Execute(findText, matchCase, matchWholeWord, false, false, false, true, 0);
                    if (!hit) break;
                    total++;
                    rng.Start = rng.End;
                    rng.End = doc.Content.End;
                    if (rng.Start >= rng.End) break;
                }
            }

            if (total == 0)
            {
                if (v2) return "已替换 0 处匹配（共找到 0 处）";
                throw new Error(string.Format("find_and_replace 失败：未找到\"{0}\"。请检查查找文本是否正确，或尝试使用 find_and_replace_v2 配合 match_case=false。", findText));
            }

            int replaced = 0;
            {
                dynamic rng = doc.Content.Duplicate;
                while (replaced < total)
                {
                    dynamic find = rng.Find;
                    find.ClearFormatting();
                    bool hit = find.Execute(findText, matchCase, matchWholeWord, false, false, false, true, 0);
                    if (!hit) break;
                    rng.Text = replaceText;
                    replaced++;
                    if (!replaceAll) break;
                    rng.Start = rng.End;
                    rng.End = doc.Content.End;
                    if (rng.Start >= rng.End) break;
                }
            }

            return v2
                ? string.Format("已替换 {0} 处匹配（共找到 {1} 处）", replaced, total)
                : string.Format("已替换 {0} 处匹配中的第 1 处", total);
        }

        private static string UndoLastAction(dynamic doc, JObject p)
        {
            int count = p["count"] != null ? (int)p["count"] : 1;
            for (int i = 0; i < count; i++)
            {
                doc.Undo();
            }
            return string.Format("已撤销最近 {0} 步操作", count);
        }

        private static string DeleteTable(dynamic doc, JObject p)
        {
            int tableIndex = p["table_index"] != null ? (int)p["table_index"] : 0;
            int tableCount = doc.Tables.Count;
            if (tableIndex < 0 || tableIndex >= tableCount)
            {
                throw new Error(string.Format("表格序号 {0} 越界（文档共有 {1} 张表格，序号范围 0 ~ {2}）。请先调用 get_document_tables 确认有效索引。",
                    tableIndex, tableCount, tableCount - 1));
            }
            doc.Tables[tableIndex + 1].Delete();
            return "已删除表格 " + tableIndex;
        }

        private static string InsertTable(dynamic doc, JObject p)
        {
            string location = (string)p["location"];
            if (string.IsNullOrEmpty(location)) throw new Error("缺少 location 参数");
            JArray rowsArr = (JArray)p["rows"];
            if (rowsArr == null || rowsArr.Count == 0) throw new Error("rows 必须是非空数组");

            var rows = new List<string[]>();
            foreach (JToken row in rowsArr)
            {
                var arr = row as JArray;
                if (arr == null) throw new Error("rows 元素必须是字符串数组");
                rows.Add(arr.Select(c => c.Type == JTokenType.Null ? "" : (string)c ?? "").ToArray());
            }
            JArray headersArr = (JArray)p["headers"];
            var allRows = new List<string[]>();
            if (headersArr != null && headersArr.Count > 0)
            {
                allRows.Add(headersArr.Select(c => c.Type == JTokenType.Null ? "" : (string)c ?? "").ToArray());
            }
            allRows.AddRange(rows);
            int colCount = allRows.Max(r => r.Length);
            if (colCount == 0) throw new Error("表格至少需要 1 列");
            var normalized = allRows.Select(r =>
            {
                var padded = r.ToList();
                while (padded.Count < colCount) padded.Add("");
                return padded.ToArray();
            }).ToList();
            int rowCount = normalized.Count;
            string style = (string)p["style"] ?? "TableGrid";

            int pos;
            string locationDesc;
            if (location == "after_heading")
            {
                string headingText = (string)p["heading_text"];
                if (string.IsNullOrEmpty(headingText)) throw new Error("location=after_heading 时必须提供 heading_text");
                int idx = FindHeadingParagraphIndex(doc, headingText);
                if (idx < 0) throw new Error("未找到标题 \"" + headingText + "\"");
                pos = ParagraphAt(doc, idx).Range.End;
                locationDesc = "已在标题 \"" + headingText + "\" 后插入";
            }
            else if (location == "after_paragraph")
            {
                int index = p["paragraph_index"] != null ? (int)p["paragraph_index"] : -1;
                if (index < 0) throw new Error("location=after_paragraph 时必须提供 paragraph_index");
                int total = (int)doc.Paragraphs.Count;
                if (index >= total) throw new Error(string.Format("paragraph_index {0} 越界（共 {1} 段）", index, total));
                pos = ParagraphAt(doc, index).Range.End;
                locationDesc = "已在段落 " + index + " 后插入";
            }
            else if (location == "at_end")
            {
                pos = (int)doc.Content.End - 1;
                locationDesc = "已追加";
            }
            else
            {
                pos = App.Selection.Range.End;
                locationDesc = "已在光标处插入";
            }

            dynamic range = doc.Range(pos, pos);
            dynamic table = doc.Tables.Add(range, rowCount, colCount);
            for (int r = 1; r <= rowCount; r++)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    table.Cell(r, c).Range.Text = normalized[r - 1][c - 1];
                }
            }
            try
            {
                table.Style = style == "TableGrid" ? "Table Grid" : style;
            }
            catch
            {
                try { table.Style = style; } catch { /* style optional */ }
            }
            return string.Format("{0} {1} × {2} 表格", locationDesc, rowCount, colCount);
        }

        private static string InsertMermaidImage(dynamic doc, JObject p)
        {
            string dataUrl = (string)p["image_data_url"];
            if (string.IsNullOrEmpty(dataUrl)) throw new Error("缺少渲染后的图片数据（image_data_url）");
            string location = (string)p["location"] ?? "at_cursor";

            int pos;
            string locationDesc;
            if (location == "after_heading")
            {
                string headingText = (string)p["heading_text"];
                if (string.IsNullOrEmpty(headingText)) throw new Error("location=after_heading 时必须提供 heading_text");
                int idx = FindHeadingParagraphIndex(doc, headingText);
                if (idx < 0) throw new Error("未找到标题 \"" + headingText + "\"");
                pos = ParagraphAt(doc, idx).Range.End;
                locationDesc = "标题 \"" + headingText + "\" 后";
            }
            else if (location == "after_paragraph")
            {
                int index = p["paragraph_index"] != null ? (int)p["paragraph_index"] : -1;
                int total = (int)doc.Paragraphs.Count;
                if (index < 0 || index >= total) throw new Error(string.Format("paragraph_index {0} 越界（共 {1} 段）", index, total));
                pos = ParagraphAt(doc, index).Range.End;
                locationDesc = "段落 " + index + " 后";
            }
            else if (location == "at_end")
            {
                pos = (int)doc.Content.End - 1;
                locationDesc = "文档末尾";
            }
            else
            {
                pos = App.Selection.Range.End;
                locationDesc = "光标处";
            }

            int comma = dataUrl.IndexOf("base64,", StringComparison.Ordinal);
            string base64 = comma >= 0 ? dataUrl.Substring(comma + 7) : dataUrl;
            string tmp = Path.Combine(Path.GetTempPath(), "oa-mermaid-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(tmp, Convert.FromBase64String(base64));
            try
            {
                dynamic range = doc.Range(pos, pos);
                doc.InlineShapes.AddPicture(tmp, false, true, range);
            }
            finally
            {
                try { File.Delete(tmp); } catch { }
            }
            return "已在" + locationDesc + "插入流程图";
        }

        // ── context providers (bridge data) ───────────────────────────────────

        private static object BuildWordContext()
        {
            dynamic doc = App.ActiveDocument;
            string bodyText = (string)doc.Content.Text ?? "";
            if (bodyText.Length > 3000) bodyText = bodyText.Substring(0, 3000);
            return new
            {
                documentContext = bodyText,
                selection = SafeText(App.Selection.Text),
            };
        }

        private static object BuildStructuredContext()
        {
            dynamic doc = App.ActiveDocument;
            int total = (int)doc.Paragraphs.Count;
            string bodyText = (string)doc.Content.Text ?? "";

            // Map table-cell paragraph indices (contiguous-run heuristic,
            // transliterated from main.ts getStructuredContext).
            var tableParagraphIndices = new HashSet<int>();
            var paragraphTableIndex = new Dictionary<int, int>();
            int tableCount = doc.Tables.Count;
            for (int t = 1; t <= tableCount; t++)
            {
                try
                {
                    dynamic table = doc.Tables[t];
                    string firstText = SafeText(table.Cell(1, 1).Range.Text).TrimEnd('\r', '\a').Trim();
                    int startIdx = -1;
                    for (int j = 0; j < total; j++)
                    {
                        if (SafeText(ParagraphAt(doc, j).Range.Text).TrimEnd('\r', '\a').Trim() == firstText) { startIdx = j; break; }
                    }
                    if (startIdx < 0) continue;
                    int runLength = table.Range.Paragraphs.Count;
                    for (int j = 0; j < runLength; j++)
                    {
                        int idx = startIdx + j;
                        if (idx >= total) break;
                        tableParagraphIndices.Add(idx);
                        paragraphTableIndex[idx] = t - 1;
                    }
                }
                catch { }
            }

            var paraList = new JArray();
            for (int i = 0; i < Math.Min(total, MaxStructuredParagraphs); i++)
            {
                dynamic para = ParagraphAt(doc, i);
                string text = SafeText(para.Range.Text).TrimEnd('\r', '\a').Trim();
                int imgCount = para.Range.InlineShapes.Count;
                if (text.Length == 0 && imgCount == 0) continue;

                string styleName = StyleNameOf(para);
                int headingLevel = HeadingLevelOf(styleName);
                dynamic font = para.Range.Font;
                string fontName = font.Name as string;
                double fontSize = (double)font.Size;
                string fontColor = ColorHex((int)font.Color);
                bool fontBold = font.Bold != 0;
                bool fontItalic = font.Italic != 0;

                bool isInTable = tableParagraphIndices.Contains(i);
                string tableMarker = isInTable ? " [表格" + paragraphTableIndex[i] + "]" : "";
                string displayText = text.Length > 0 ? text : (imgCount > 0 ? "（图片段落，无文字）" : "");
                if (displayText.Length > MaxStructuredText) displayText = displayText.Substring(0, MaxStructuredText);

                var item = new JObject
                {
                    ["index"] = i,
                    ["text"] = displayText + (imgCount > 0 ? " [📷图片×" + imgCount + "]" : "") + tableMarker,
                    ["style"] = styleName,
                    ["isTable"] = isInTable,
                    ["isList"] = para.Range.ListFormat.ListType != 0,
                };
                if (headingLevel > 0) item["headingLevel"] = headingLevel;
                var fontObj = new JObject();
                if (!string.IsNullOrEmpty(fontName)) fontObj["name"] = fontName;
                if (fontSize > 0) fontObj["size"] = fontSize;
                if (fontColor != null) fontObj["color"] = fontColor;
                if (fontBold) fontObj["bold"] = true;
                if (fontItalic) fontObj["italic"] = true;
                if (fontObj.Count > 0) item["font"] = fontObj;
                paraList.Add(item);
            }

            dynamic sel = App.Selection;
            string selText = SafeText(sel.Text);
            int? startIdx2 = null, endIdx2 = null;
            if (selText.Length > 0)
            {
                string firstSel = SafeText(sel.Paragraphs[1].Range.Text).TrimEnd('\r', '\a').Trim();
                string lastSel = SafeText(sel.Paragraphs[sel.Paragraphs.Count].Range.Text).TrimEnd('\r', '\a').Trim();
                foreach (JObject item in paraList)
                {
                    string pText = (string)item["text"] ?? "";
                    if (startIdx2 == null && (pText == firstSel || (firstSel.Length > 0 && pText.StartsWith(firstSel)) || (pText.Length > 0 && firstSel.StartsWith(pText))))
                        startIdx2 = (int)item["index"];
                    if (endIdx2 == null && (pText == lastSel || (lastSel.Length > 0 && pText.StartsWith(lastSel)) || (pText.Length > 0 && lastSel.StartsWith(pText))))
                        endIdx2 = (int)item["index"];
                }
            }

            var selection = new JObject { ["text"] = selText.Length > 2000 ? selText.Substring(0, 2000) : selText };
            if (startIdx2 != null) selection["startParagraphIndex"] = startIdx2.Value;
            if (endIdx2 != null) selection["endParagraphIndex"] = endIdx2.Value;

            return new
            {
                title = "",
                totalParagraphs = total,
                totalCharacters = bodyText.Length,
                paragraphs = paraList,
                selection,
            };
        }

        // ── simple insert (preview flow) ──────────────────────────────────────

        private static ToolResult SimpleInsert(string mode, string text)
        {
            try
            {
                dynamic doc = App.ActiveDocument;
                switch (mode)
                {
                    case "replace":
                        App.Selection.Range.Text = text;
                        return new ToolResult { Success = true, Result = "已替换当前选区" };
                    case "append_end":
                    {
                        int pos = (int)doc.Content.End - 1;
                        dynamic r = doc.Range(pos, pos);
                        r.InsertAfter(text + "\r");
                        return new ToolResult { Success = true, Result = "已追加到文档末尾" };
                    }
                    case "insert_start":
                    {
                        int pos = doc.Paragraphs[1].Range.Start;
                        dynamic r = doc.Range(pos, pos);
                        r.InsertAfter(text + "\r");
                        return new ToolResult { Success = true, Result = "已插入到文档开头" };
                    }
                    case "at_cursor":
                    {
                        App.Selection.Range.InsertAfter(text);
                        return new ToolResult { Success = true, Result = "已在光标处插入" };
                    }
                    default:
                        return new ToolResult { Success = false, Error = "未知插入模式: " + mode };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Success = false, Error = ex.Message };
            }
        }

        // ── shared helpers ────────────────────────────────────────────────────

        private static dynamic ParagraphAt(dynamic doc, int zeroBasedIndex)
        {
            return doc.Paragraphs[zeroBasedIndex + 1];
        }

        /// <summary>
        /// COM properties with optional parameters (e.g. Range.Runs) cannot be
        /// bound by the C# dynamic binder — invoke them via IDispatch instead.
        /// </summary>
        private static object GetComObjectProperty(object target, string name)
        {
            return target.GetType().InvokeMember(
                name, System.Reflection.BindingFlags.GetProperty, null, target, null);
        }

        private static string SafeText(object text)
        {
            return text as string ?? "";
        }

        private static string StyleNameOf(dynamic paragraphOrTable)
        {
            try
            {
                dynamic style = paragraphOrTable.Range.Style;
                string nameLocal = style.NameLocal as string;
                return string.IsNullOrEmpty(nameLocal) ? "Normal" : nameLocal;
            }
            catch
            {
                try { return (string)paragraphOrTable.Style ?? "Normal"; } catch { return "Normal"; }
            }
        }

        private static int HeadingLevelOf(string styleName)
        {
            if (string.IsNullOrEmpty(styleName)) return 0;
            Match m = Regex.Match(styleName, @"(?:heading|标题)\s*(\d)", RegexOptions.IgnoreCase);
            return m.Success ? int.Parse(m.Groups[1].Value) : 0;
        }

        private static int FindHeadingParagraphIndex(dynamic doc, string headingText)
        {
            // Faithful to the TS implementation: body.search() returns the first
            // paragraph containing the text regardless of style.
            int total = (int)doc.Paragraphs.Count;
            for (int i = 0; i < total; i++)
            {
                string text = SafeText(ParagraphAt(doc, i).Range.Text).TrimEnd('\r', '\a').Trim();
                if (text.Length > 0 && text.IndexOf(headingText, StringComparison.OrdinalIgnoreCase) >= 0) return i;
            }
            return -1;
        }

        private static string[] NormalizeContentLines(string content)
        {
            string text = (content ?? "").Replace("\\r\\n", "\n").Replace("\\n", "\n");
            text = Regex.Replace(text, "\n{3,}", "\n\n");
            var lines = text.Split('\n').ToList();
            if (lines.Count > 1 && lines[lines.Count - 1].Trim().Length == 0) lines.RemoveAt(lines.Count - 1);
            return lines.ToArray();
        }

        private static int InsertMultiParagraphAt(dynamic doc, int pos, string content, string format)
        {
            string[] lines = NormalizeContentLines(content);
            dynamic r = doc.Range(pos, pos);
            r.InsertAfter(string.Join("\r", lines) + "\r");
            ApplyFormatToRange(r, format);
            return lines.Length;
        }

        private static void ApplyFormatToRange(dynamic range, string format)
        {
            if (string.IsNullOrEmpty(format)) return;
            if (!StyleMap.ContainsKey(format)) return;
            for (int i = 1; i <= range.Paragraphs.Count; i++)
            {
                ApplyFormatToParagraph(range.Paragraphs[i], format);
            }
        }

        private static void ApplyFormatToParagraph(dynamic para, string format)
        {
            if (string.IsNullOrEmpty(format) || !StyleMap.ContainsKey(format)) return;
            foreach (object candidate in StyleMap[format])
            {
                try
                {
                    para.Style = candidate;
                    return;
                }
                catch
                {
                    // try next candidate
                }
            }
            Addin.Trace("WordToolExecutor: 无法设置样式 " + format);
        }

        private static void SetParagraphText(dynamic para, string text)
        {
            dynamic r = para.Range.Duplicate;
            r.End = r.End - 1; // keep the paragraph mark
            r.Text = text;
        }

        private static void InsertHtmlAt(dynamic doc, int pos, string html)
        {
            dynamic r = doc.Range(pos, pos);
            InsertHtmlAtRange(r, html);
        }

        private static void InsertHtmlAtRange(dynamic range, string html)
        {
            string tmp = Path.Combine(Path.GetTempPath(), "oa-insert-" + Guid.NewGuid().ToString("N") + ".html");
            File.WriteAllText(tmp, html, Encoding.UTF8);
            try
            {
                range.InsertFile(tmp);
            }
            finally
            {
                try { File.Delete(tmp); } catch { }
            }
        }

        private static string ColorHex(int wdColor)
        {
            if (wdColor == -16777216 || wdColor < 0) return null; // wdColorAutomatic / theme
            int r = wdColor & 0xFF;
            int g = (wdColor >> 8) & 0xFF;
            int b = (wdColor >> 16) & 0xFF;
            return string.Format("#{0:X2}{1:X2}{2:X2}", r, g, b);
        }

        private static int WdColorFromHex(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return -16777216;
            string h = hex.TrimStart('#');
            if (h.Length == 8) h = h.Substring(0, 6);
            if (h.Length != 6) return -16777216;
            int r = Convert.ToInt32(h.Substring(0, 2), 16);
            int g = Convert.ToInt32(h.Substring(2, 2), 16);
            int b = Convert.ToInt32(h.Substring(4, 2), 16);
            return r + g * 256 + b * 65536;
        }

        internal sealed class Error : Exception
        {
            public Error(string message) : base(message) { }
        }
    }
}
