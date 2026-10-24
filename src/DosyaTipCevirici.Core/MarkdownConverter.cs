using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace DosyaTipCevirici.Core;

public static partial class MarkdownConverter
{
    private static readonly Regex FrontMatterRegex = new(@"\A---\r?\n.*?\r?\n---\r?\n", RegexOptions.Singleline);
    private static readonly Regex CodeBlockRegex = new(@"^(?:```|~~~)(?<lang>[^\r\n]*)\r?\n(?<code>.*?)\r?\n(?:```|~~~)$", RegexOptions.Multiline | RegexOptions.Singleline);
    private static readonly Regex InlineCodeRegex = new(@"`([^`\r\n]+)`");
    private static readonly Regex ImageRegex = new(@"!\[(?<alt>[^\]]*)\]\((?<url>[^)\s]+)(?:\s+""(?<title>[^""]*)"")?\)");
    private static readonly Regex LinkRegex = new(@"(?<!!)\[(?<text>[^\]]+)\]\((?<url>[^)\s]+)(?:\s+""(?<title>[^""]*)"")?\)");
    private static readonly Regex BoldItalicRegex = new(@"(?:\*\*\*|___)(?=\S)(.+?)(?<=\S)(?:\*\*\*|___)");
    private static readonly Regex BoldRegex = new(@"(?:\*\*|__)(?=\S)(.+?)(?<=\S)(?:\*\*|__)");
    private static readonly Regex ItalicRegex = new(@"(?:\*|_)(?=\S)(.+?)(?<=\S)(?:\*|_)");
    private static readonly Regex StrikethroughRegex = new(@"~~(?=\S)(.+?)(?<=\S)~~");
    private static readonly Regex AutolinkRegex = new(@"<((?:https?|mailto):[^>]+)>");

    public static string ToHtml(string markdown, string title = "Belge", bool fullDocument = true)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return fullDocument ? WrapHtml("", title) : "";

        var text = FrontMatterRegex.Replace(markdown, "");
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var sb = new StringBuilder();

        var inCodeBlock = false;
        var codeLang = "";
        var codeLines = new List<string>();

        var inList = false;
        var listStack = new Stack<(bool IsOrdered, int Indent)>();

        var inBlockquote = false;
        var blockquoteLines = new List<string>();

        var inTable = false;
        var tableLines = new List<string>();

        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var trimmed = line.Trim();

            // Code blocks
            if (!inCodeBlock && (trimmed.StartsWith("```") || trimmed.StartsWith("~~~")))
            {
                CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines);
                inCodeBlock = true;
                codeLang = trimmed[3..].Trim();
                codeLines.Clear();
                continue;
            }

            if (inCodeBlock)
            {
                if (trimmed.StartsWith("```") || trimmed.StartsWith("~~~"))
                {
                    inCodeBlock = false;
                    var langAttr = string.IsNullOrWhiteSpace(codeLang) ? "" : $" class=\"language-{HtmlEscape(codeLang)}\"";
                    sb.Append("<pre><code").Append(langAttr).Append('>')
                      .Append(HtmlEscape(string.Join("\n", codeLines)))
                      .AppendLine("</code></pre>");
                }
                else
                {
                    codeLines.Add(line);
                }
                continue;
            }

            // Tables: line containing |
            if (trimmed.StartsWith('|') && trimmed.EndsWith('|') && trimmed.Length >= 2)
            {
                CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines, keepTable: true);
                inTable = true;
                tableLines.Add(trimmed);
                continue;
            }
            else if (inTable)
            {
                FlushTable(sb, tableLines);
                inTable = false;
            }

            // Blockquotes
            if (trimmed.StartsWith('>'))
            {
                CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines, keepBlockquote: true);
                inBlockquote = true;
                var quoteContent = trimmed[1..].TrimStart();
                blockquoteLines.Add(quoteContent);
                continue;
            }
            else if (inBlockquote)
            {
                FlushBlockquote(sb, blockquoteLines);
                inBlockquote = false;
            }

            // Horizontal Rule
            if (IsHorizontalRule(trimmed))
            {
                CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines);
                sb.AppendLine("<hr />");
                continue;
            }

            // Setext Headings (underlined with === or ---)
            if (i + 1 < lines.Length && !string.IsNullOrWhiteSpace(trimmed))
            {
                var nextTrimmed = lines[i + 1].Trim();
                if (nextTrimmed.Length >= 3 && nextTrimmed.All(c => c == '='))
                {
                    CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines);
                    sb.Append("<h1>").Append(FormatInline(trimmed)).AppendLine("</h1>");
                    i++;
                    continue;
                }
                if (nextTrimmed.Length >= 3 && nextTrimmed.All(c => c == '-') && !trimmed.StartsWith('-') && !trimmed.StartsWith('*'))
                {
                    CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines);
                    sb.Append("<h2>").Append(FormatInline(trimmed)).AppendLine("</h2>");
                    i++;
                    continue;
                }
            }

            // ATX Headings (# Heading)
            if (trimmed.StartsWith('#'))
            {
                var level = 0;
                while (level < trimmed.Length && trimmed[level] == '#') level++;
                if (level is >= 1 and <= 6 && (level == trimmed.Length || trimmed[level] == ' '))
                {
                    CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines);
                    var headingText = trimmed[level..].Trim();
                    sb.Append($"<h{level}>").Append(FormatInline(headingText)).AppendLine($"</h{level}>");
                    continue;
                }
            }

            // Lists (Ordered & Unordered)
            if (IsListItem(line, out var isOrdered, out var indent, out var content))
            {
                if (inTable) { FlushTable(sb, tableLines); inTable = false; }
                if (inBlockquote) { FlushBlockquote(sb, blockquoteLines); inBlockquote = false; }

                AdjustListLevel(sb, ref inList, listStack, isOrdered, indent);

                // Task list item check: [ ] or [x]
                string itemHtml;
                if (content.StartsWith("[ ] "))
                    itemHtml = "<input type=\"checkbox\" disabled /> " + FormatInline(content[4..]);
                else if (content.StartsWith("[x] ", StringComparison.OrdinalIgnoreCase))
                    itemHtml = "<input type=\"checkbox\" checked disabled /> " + FormatInline(content[4..]);
                else
                    itemHtml = FormatInline(content);

                sb.Append("<li>").Append(itemHtml).AppendLine("</li>");
                continue;
            }
            else if (inList && string.IsNullOrWhiteSpace(trimmed))
            {
                // Blank line might end list or separate items
                CloseList(sb, ref inList, listStack);
                continue;
            }

            // Blank line
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines);
                continue;
            }

            // Regular paragraph
            CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines);

            // Collect multiline paragraph
            var paraLines = new List<string> { trimmed };
            while (i + 1 < lines.Length && !string.IsNullOrWhiteSpace(lines[i + 1].Trim())
                   && !lines[i + 1].Trim().StartsWith('#')
                   && !lines[i + 1].Trim().StartsWith('>')
                   && !lines[i + 1].Trim().StartsWith("```")
                   && !lines[i + 1].Trim().StartsWith("~~~")
                   && !IsHorizontalRule(lines[i + 1].Trim())
                   && !IsListItem(lines[i + 1], out _, out _, out _)
                   && !(lines[i + 1].Trim().StartsWith('|') && lines[i + 1].Trim().EndsWith('|')))
            {
                i++;
                paraLines.Add(lines[i].Trim());
            }

            var paraText = string.Join(" ", paraLines);
            sb.Append("<p>").Append(FormatInline(paraText)).AppendLine("</p>");
        }

        // Close any remaining open blocks
        if (inCodeBlock)
        {
            sb.Append("<pre><code>")
              .Append(HtmlEscape(string.Join("\n", codeLines)))
              .AppendLine("</code></pre>");
        }
        CloseContainers(sb, ref inList, listStack, ref inBlockquote, blockquoteLines, ref inTable, tableLines);

        var body = sb.ToString();
        return fullDocument ? WrapHtml(body, title) : body;
    }

    public static string FromHtml(string html)
    {
        if (string.IsNullOrWhiteSpace(html)) return "";

        // Remove script, style, comments, and head
        var cleaned = Regex.Replace(html, @"<!--.*?-->", "", RegexOptions.Singleline);
        cleaned = Regex.Replace(cleaned, @"<head\b[^>]*>.*?</head>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        cleaned = Regex.Replace(cleaned, @"<style\b[^>]*>.*?</style>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        cleaned = Regex.Replace(cleaned, @"<script\b[^>]*>.*?</script>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Pre / code blocks: extract them first to protect from formatting
        var codeBlocks = new List<string>();
        cleaned = Regex.Replace(cleaned, @"<pre\b[^>]*>(?:<code(?:\s+class=""language-(?<lang>[^""]+)"")?[^>]*>)?(?<code>.*?)(?:</code>)?</pre>", match =>
        {
            var code = WebUtility.HtmlDecode(match.Groups["code"].Value);
            var lang = match.Groups["lang"].Value;
            var placeholder = $"__CODE_BLOCK_{codeBlocks.Count}__";
            codeBlocks.Add($"```{lang}\n{code.TrimEnd()}\n```\n\n");
            return placeholder;
        }, RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Tables
        cleaned = Regex.Replace(cleaned, @"<table\b[^>]*>(.*?)</table>", match =>
        {
            var tableContent = match.Groups[1].Value;
            var rows = Regex.Matches(tableContent, @"<tr\b[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (rows.Count == 0) return "";

            var tableSb = new StringBuilder();
            var isHeader = true;
            var colCount = 0;

            foreach (Match row in rows)
            {
                var cells = Regex.Matches(row.Groups[1].Value, @"<(?:th|td)\b[^>]*>(.*?)</(?:th|td)>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                if (cells.Count == 0) continue;

                var cellValues = cells.Select(c => StripHtml(c.Groups[1].Value).Replace("\n", " ").Trim()).ToArray();
                tableSb.Append("| ").Append(string.Join(" | ", cellValues)).AppendLine(" |");

                if (isHeader)
                {
                    colCount = cellValues.Length;
                    tableSb.Append("| ").Append(string.Join(" | ", Enumerable.Repeat("---", colCount))).AppendLine(" |");
                    isHeader = false;
                }
            }
            return tableSb.AppendLine().ToString();
        }, RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Headings
        for (var h = 1; h <= 6; h++)
        {
            var prefix = new string('#', h) + " ";
            cleaned = Regex.Replace(cleaned, $@"<h{h}\b[^>]*>(.*?)</h{h}>", m => prefix + CleanInlineHtml(m.Groups[1].Value).Trim() + "\n\n", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        }

        // Horizontal Rules
        cleaned = Regex.Replace(cleaned, @"<hr\b[^>]*>", "---\n\n", RegexOptions.IgnoreCase);

        // Blockquotes
        cleaned = Regex.Replace(cleaned, @"<blockquote\b[^>]*>(.*?)</blockquote>", m =>
        {
            var inner = CleanInlineHtml(m.Groups[1].Value).Trim();
            var lines = inner.Split('\n').Select(l => "> " + l);
            return string.Join("\n", lines) + "\n\n";
        }, RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Lists
        cleaned = Regex.Replace(cleaned, @"<ul\b[^>]*>(.*?)</ul>", m =>
        {
            var items = Regex.Matches(m.Groups[1].Value, @"<li\b[^>]*>(.*?)</li>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            var sb = new StringBuilder();
            foreach (Match item in items)
                sb.Append("- ").AppendLine(CleanInlineHtml(item.Groups[1].Value).Trim());
            return sb.AppendLine().ToString();
        }, RegexOptions.Singleline | RegexOptions.IgnoreCase);

        cleaned = Regex.Replace(cleaned, @"<ol\b[^>]*>(.*?)</ol>", m =>
        {
            var items = Regex.Matches(m.Groups[1].Value, @"<li\b[^>]*>(.*?)</li>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            var sb = new StringBuilder();
            var num = 1;
            foreach (Match item in items)
                sb.Append($"{num++}. ").AppendLine(CleanInlineHtml(item.Groups[1].Value).Trim());
            return sb.AppendLine().ToString();
        }, RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Paragraphs and Divs
        cleaned = Regex.Replace(cleaned, @"<(?:p|div)\b[^>]*>(.*?)</(?:p|div)>", m => CleanInlineHtml(m.Groups[1].Value).Trim() + "\n\n", RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Line breaks
        cleaned = Regex.Replace(cleaned, @"<br\b[^>]*>", "\n", RegexOptions.IgnoreCase);

        // Final inline cleanup on remainder
        cleaned = CleanInlineHtml(cleaned);

        // Restore code blocks
        for (var i = 0; i < codeBlocks.Count; i++)
            cleaned = cleaned.Replace($"__CODE_BLOCK_{i}__", codeBlocks[i]);

        // Clean up whitespace: max 2 consecutive newlines, trim lines
        var resultLines = cleaned.Split('\n').Select(l => l.TrimEnd());
        var result = string.Join("\n", resultLines);
        result = Regex.Replace(result, @"\n{3,}", "\n\n");
        return result.Trim() + "\n";
    }

    public static string ToPlainText(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown)) return "";
        var text = FrontMatterRegex.Replace(markdown, "");

        // Remove code block markers
        text = Regex.Replace(text, @"^(?:```|~~~)[^\r\n]*", "", RegexOptions.Multiline);
        text = Regex.Replace(text, @"(?:```|~~~)$", "", RegexOptions.Multiline);

        // Remove headers
        text = Regex.Replace(text, @"^#{1,6}\s*", "", RegexOptions.Multiline);

        // Links: [text](url) -> text
        text = LinkRegex.Replace(text, "${text}");
        // Images: ![alt](url) -> alt
        text = ImageRegex.Replace(text, "${alt}");
        // Inline formatting
        text = BoldItalicRegex.Replace(text, "$1");
        text = BoldRegex.Replace(text, "$1");
        text = ItalicRegex.Replace(text, "$1");
        text = StrikethroughRegex.Replace(text, "$1");
        text = InlineCodeRegex.Replace(text, "$1");
        text = AutolinkRegex.Replace(text, "$1");

        // Clean tables: replace pipes with tab or space
        text = Regex.Replace(text, @"^\|[-:|\s]+\|$", "", RegexOptions.Multiline);
        text = Regex.Replace(text, @"\|", "\t");

        return text.Trim();
    }

    private static string CleanInlineHtml(string html)
    {
        if (string.IsNullOrEmpty(html)) return "";

        // Links
        var text = Regex.Replace(html, @"<a\b[^>]*href=[""'](?<url>[^""']+)[""'][^>]*>(?<text>.*?)</a>", m =>
        {
            var linkText = StripHtml(m.Groups["text"].Value).Trim();
            var url = m.Groups["url"].Value.Trim();
            return string.IsNullOrEmpty(linkText) ? url : $"[{linkText}]({url})";
        }, RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Images
        text = Regex.Replace(text, @"<img\b[^>]*src=[""'](?<src>[^""']+)[""'](?:[^>]*alt=[""'](?<alt>[^""']*)[""'])?[^>]*>", m =>
        {
            var src = m.Groups["src"].Value;
            var alt = m.Groups["alt"].Value;
            return $"![{alt}]({src})";
        }, RegexOptions.IgnoreCase);

        // Bold & Italic
        text = Regex.Replace(text, @"<(?:b|strong)\b[^>]*>(.*?)</(?:b|strong)>", m => $"**{StripHtml(m.Groups[1].Value)}**", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"<(?:i|em)\b[^>]*>(.*?)</(?:i|em)>", m => $"*{StripHtml(m.Groups[1].Value)}*", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"<(?:del|s|strike)\b[^>]*>(.*?)</(?:del|s|strike)>", m => $"~~{StripHtml(m.Groups[1].Value)}~~", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"<code\b[^>]*>(.*?)</code>", m => $"`{StripHtml(m.Groups[1].Value)}`", RegexOptions.Singleline | RegexOptions.IgnoreCase);

        // Strip remaining HTML tags
        text = StripHtml(text);

        // Decode entities
        return WebUtility.HtmlDecode(text);
    }

    private static string StripHtml(string input) =>
        Regex.Replace(input, @"<[^>]+>", "");

    private static string HtmlEscape(string text) =>
        string.IsNullOrEmpty(text) ? "" :
        text.Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");

    private static string FormatInline(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";

        // First escape HTML entities
        var encoded = HtmlEscape(text);

        // Inline code `` `code` ``
        encoded = InlineCodeRegex.Replace(encoded, match =>
            $"<code>{match.Groups[1].Value}</code>");

        // Images
        encoded = ImageRegex.Replace(encoded, match =>
        {
            var alt = match.Groups["alt"].Value;
            var url = match.Groups["url"].Value;
            var title = match.Groups["title"].Value;
            var titleAttr = string.IsNullOrEmpty(title) ? "" : $" title=\"{title}\"";
            return $"<img src=\"{url}\" alt=\"{alt}\"{titleAttr} />";
        });

        // Links
        encoded = LinkRegex.Replace(encoded, match =>
        {
            var linkText = match.Groups["text"].Value;
            var url = match.Groups["url"].Value;
            var title = match.Groups["title"].Value;
            var titleAttr = string.IsNullOrEmpty(title) ? "" : $" title=\"{title}\"";
            return $"<a href=\"{url}\"{titleAttr}>{linkText}</a>";
        });

        // Autolinks <http://...>
        encoded = AutolinkRegex.Replace(encoded, match =>
        {
            var url = match.Groups[1].Value;
            return $"<a href=\"{url}\">{url}</a>";
        });

        // Bold & Italic
        encoded = BoldItalicRegex.Replace(encoded, "<strong><em>$1</em></strong>");
        encoded = BoldRegex.Replace(encoded, "<strong>$1</strong>");
        encoded = ItalicRegex.Replace(encoded, "<em>$1</em>");
        encoded = StrikethroughRegex.Replace(encoded, "<del>$1</del>");

        return encoded;
    }

    private static bool IsHorizontalRule(string line)
    {
        if (line.Length < 3) return false;
        var stripped = line.Replace(" ", "");
        return stripped.Length >= 3 &&
               (stripped.All(c => c == '-') || stripped.All(c => c == '*') || stripped.All(c => c == '_'));
    }

    private static bool IsListItem(string line, out bool isOrdered, out int indent, out string content)
    {
        isOrdered = false;
        indent = 0;
        content = "";

        if (string.IsNullOrWhiteSpace(line)) return false;

        indent = line.TakeWhile(char.IsWhiteSpace).Count();
        var trimmed = line.TrimStart();

        // Unordered list: - , * , +
        if ((trimmed.StartsWith("- ") || trimmed.StartsWith("* ") || trimmed.StartsWith("+ ")) && trimmed.Length >= 2)
        {
            isOrdered = false;
            content = trimmed[2..].Trim();
            return true;
        }

        // Ordered list: 1. , 2. 
        var dotIndex = trimmed.IndexOf(". ", StringComparison.Ordinal);
        if (dotIndex > 0 && int.TryParse(trimmed[..dotIndex], out _))
        {
            isOrdered = true;
            content = trimmed[(dotIndex + 2)..].Trim();
            return true;
        }

        return false;
    }

    private static void AdjustListLevel(StringBuilder sb, ref bool inList, Stack<(bool IsOrdered, int Indent)> stack, bool isOrdered, int indent)
    {
        if (!inList)
        {
            inList = true;
            stack.Push((isOrdered, indent));
            sb.AppendLine(isOrdered ? "<ol>" : "<ul>");
            return;
        }

        var current = stack.Peek();
        if (indent > current.Indent)
        {
            stack.Push((isOrdered, indent));
            sb.AppendLine(isOrdered ? "<ol>" : "<ul>");
        }
        else if (indent < current.Indent)
        {
            while (stack.Count > 0 && indent < stack.Peek().Indent)
            {
                var popped = stack.Pop();
                sb.AppendLine(popped.IsOrdered ? "</ol>" : "</ul>");
            }

            if (stack.Count == 0 || stack.Peek().IsOrdered != isOrdered)
            {
                if (stack.Count > 0)
                {
                    var popped = stack.Pop();
                    sb.AppendLine(popped.IsOrdered ? "</ol>" : "</ul>");
                }
                stack.Push((isOrdered, indent));
                sb.AppendLine(isOrdered ? "<ol>" : "<ul>");
            }
        }
        else if (current.IsOrdered != isOrdered)
        {
            var popped = stack.Pop();
            sb.AppendLine(popped.IsOrdered ? "</ol>" : "</ul>");
            stack.Push((isOrdered, indent));
            sb.AppendLine(isOrdered ? "<ol>" : "<ul>");
        }
    }

    private static void CloseList(StringBuilder sb, ref bool inList, Stack<(bool IsOrdered, int Indent)> stack)
    {
        while (stack.Count > 0)
        {
            var popped = stack.Pop();
            sb.AppendLine(popped.IsOrdered ? "</ol>" : "</ul>");
        }
        inList = false;
    }

    private static void FlushTable(StringBuilder sb, List<string> lines)
    {
        if (lines.Count < 2)
        {
            foreach (var l in lines) sb.Append("<p>").Append(FormatInline(l)).AppendLine("</p>");
            lines.Clear();
            return;
        }

        var headerRow = ParseTableRow(lines[0]);
        var sepRow = ParseTableRow(lines[1]);
        if (sepRow.All(c => c.All(ch => ch is '-' or ':' or ' ')))
        {
            var alignments = sepRow.Select(col =>
            {
                var t = col.Trim();
                if (t.StartsWith(':') && t.EndsWith(':')) return "center";
                if (t.EndsWith(':')) return "right";
                return "left";
            }).ToArray();

            sb.AppendLine("<table>");
            sb.AppendLine("<thead><tr>");
            for (var i = 0; i < headerRow.Length; i++)
            {
                var align = i < alignments.Length ? alignments[i] : "left";
                sb.Append($"<th style=\"text-align:{align}\">").Append(FormatInline(headerRow[i])).AppendLine("</th>");
            }
            sb.AppendLine("</tr></thead>");

            sb.AppendLine("<tbody>");
            for (var r = 2; r < lines.Count; r++)
            {
                var rowCells = ParseTableRow(lines[r]);
                sb.AppendLine("<tr>");
                for (var i = 0; i < headerRow.Length; i++)
                {
                    var text = i < rowCells.Length ? rowCells[i] : "";
                    var align = i < alignments.Length ? alignments[i] : "left";
                    sb.Append($"<td style=\"text-align:{align}\">").Append(FormatInline(text)).AppendLine("</td>");
                }
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody></table>");
        }
        else
        {
            foreach (var l in lines) sb.Append("<p>").Append(FormatInline(l)).AppendLine("</p>");
        }
        lines.Clear();
    }

    private static string[] ParseTableRow(string row)
    {
        var trimmed = row.Trim();
        if (trimmed.StartsWith('|')) trimmed = trimmed[1..];
        if (trimmed.EndsWith('|')) trimmed = trimmed[..^1];
        return trimmed.Split('|').Select(c => c.Trim()).ToArray();
    }

    private static void FlushBlockquote(StringBuilder sb, List<string> lines)
    {
        if (lines.Count == 0) return;
        sb.AppendLine("<blockquote>");
        foreach (var l in lines)
            sb.Append("<p>").Append(FormatInline(l)).AppendLine("</p>");
        sb.AppendLine("</blockquote>");
        lines.Clear();
    }

    private static void CloseContainers(StringBuilder sb, ref bool inList, Stack<(bool IsOrdered, int Indent)> listStack,
        ref bool inBlockquote, List<string> blockquoteLines, ref bool inTable, List<string> tableLines,
        bool keepList = false, bool keepBlockquote = false, bool keepTable = false)
    {
        if (!keepList && inList) CloseList(sb, ref inList, listStack);
        if (!keepBlockquote && inBlockquote) { FlushBlockquote(sb, blockquoteLines); inBlockquote = false; }
        if (!keepTable && inTable) { FlushTable(sb, tableLines); inTable = false; }
    }

    private static string WrapHtml(string body, string title) =>
        $$"""
        <!DOCTYPE html>
        <html lang="tr">
        <head>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>{{HtmlEscape(title)}}</title>
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
              line-height: 1.6;
              color: #1a202c;
              background-color: #ffffff;
              max-width: 860px;
              margin: 40px auto;
              padding: 0 24px;
            }
            h1, h2, h3, h4, h5, h6 { color: #1a365d; margin-top: 1.5em; margin-bottom: 0.5em; font-weight: 600; }
            h1 { font-size: 2em; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.3em; }
            h2 { font-size: 1.5em; border-bottom: 1px solid #edf2f7; padding-bottom: 0.25em; }
            p, ul, ol, table, blockquote, pre { margin-bottom: 1em; }
            code { font-family: "Consolas", "Courier New", monospace; background-color: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 0.9em; }
            pre { background-color: #0f172a; color: #f8fafc; padding: 16px; border-radius: 8px; overflow-x: auto; font-family: "Consolas", "Courier New", monospace; }
            pre code { background: none; color: inherit; padding: 0; }
            blockquote { border-left: 4px solid #3b82f6; margin-left: 0; padding-left: 16px; color: #475569; font-style: italic; background: #f8fafc; padding-top: 8px; padding-bottom: 8px; }
            table { border-collapse: collapse; width: 100%; margin: 16px 0; }
            th, td { border: 1px solid #cbd5e1; padding: 8px 12px; text-align: left; }
            th { background-color: #f1f5f9; font-weight: 600; }
            tr:nth-child(even) td { background-color: #f8fafc; }
            hr { border: 0; height: 1px; background: #e2e8f0; margin: 24px 0; }
            a { color: #2563eb; text-decoration: none; }
            a:hover { text-decoration: underline; }
            input[type="checkbox"] { margin-right: 6px; }
          </style>
        </head>
        <body>
        {{body}}
        </body>
        </html>
        """;
}
