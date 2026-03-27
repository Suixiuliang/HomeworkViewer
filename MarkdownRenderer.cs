#nullable disable
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;

namespace HomeworkViewer
{
    public class FormattedText
    {
        public string Text { get; set; }
        public FontStyle Style { get; set; }
        public Color? Color { get; set; }
        public string Url { get; set; }
        public bool IsLatex { get; set; }
    }

    public enum ParagraphType
    {
        Normal,
        Heading,
        Code,
        Quote,
        ListItem,
        Table
    }

    public class Paragraph
    {
        public List<FormattedText> FormattedParts { get; set; } = new List<FormattedText>();
        public ParagraphType Type { get; set; } = ParagraphType.Normal;
        public int HeadingLevel { get; set; } = 0;
        public bool IsListItem { get; set; }
        public int ListItemNumber { get; set; }
        public bool IsOrderedList { get; set; }
    }

    public class MarkdownRenderer
    {
        private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .Build();

        // LaTeX 命令映射（完整版）
        private static readonly Dictionary<string, string> LatexToUnicode = new Dictionary<string, string>
        {
            { "\\alpha", "α" }, { "\\beta", "β" }, { "\\gamma", "γ" }, { "\\delta", "δ" },
            { "\\epsilon", "ε" }, { "\\varepsilon", "ε" }, { "\\zeta", "ζ" }, { "\\eta", "η" },
            { "\\theta", "θ" }, { "\\iota", "ι" }, { "\\kappa", "κ" }, { "\\lambda", "λ" },
            { "\\mu", "μ" }, { "\\nu", "ν" }, { "\\xi", "ξ" }, { "\\omicron", "ο" },
            { "\\pi", "π" }, { "\\rho", "ρ" }, { "\\sigma", "σ" }, { "\\tau", "τ" },
            { "\\upsilon", "υ" }, { "\\phi", "φ" }, { "\\chi", "χ" }, { "\\psi", "ψ" },
            { "\\omega", "ω" },
            { "\\Gamma", "Γ" }, { "\\Delta", "Δ" }, { "\\Theta", "Θ" }, { "\\Lambda", "Λ" },
            { "\\Xi", "Ξ" }, { "\\Pi", "Π" }, { "\\Sigma", "Σ" }, { "\\Upsilon", "Υ" },
            { "\\Phi", "Φ" }, { "\\Psi", "Ψ" }, { "\\Omega", "Ω" },
            { "\\infty", "∞" }, { "\\pm", "±" }, { "\\mp", "∓" }, { "\\times", "×" },
            { "\\div", "÷" }, { "\\cdot", "·" }, { "\\sqrt", "√" }, { "\\int", "∫" },
            { "\\sum", "∑" }, { "\\prod", "∏" }, { "\\partial", "∂" }, { "\\nabla", "∇" },
            { "\\angle", "∠" }, { "\\perp", "⊥" }, { "\\parallel", "∥" }, { "\\approx", "≈" },
            { "\\neq", "≠" }, { "\\leq", "≤" }, { "\\geq", "≥" }, { "\\subset", "⊂" },
            { "\\supset", "⊃" }, { "\\subseteq", "⊆" }, { "\\supseteq", "⊇" }, { "\\in", "∈" },
            { "\\notin", "∉" }, { "\\cup", "∪" }, { "\\cap", "∩" }, { "\\emptyset", "∅" },
            { "\\forall", "∀" }, { "\\exists", "∃" }, { "\\neg", "¬" }, { "\\to", "→" },
            { "\\leftarrow", "←" }, { "\\rightarrow", "→" }, { "\\uparrow", "↑" }, { "\\downarrow", "↓" },
            { "\\prime", "′" }, { "\\degree", "°" }, { "\\circ", "∘" }, { "\\ldots", "…" },
            { "\\cdots", "⋯" }, { "\\vdots", "⋮" }, { "\\ddots", "⋱" },
            { "\\hat", "^" }, { "\\bar", "¯" }, { "\\vec", "→" }, { "\\dot", "˙" }
        };

        private static readonly Dictionary<char, char> SuperscriptMap = new Dictionary<char, char>
        {
            { '0', '⁰' }, { '1', '¹' }, { '2', '²' }, { '3', '³' }, { '4', '⁴' },
            { '5', '⁵' }, { '6', '⁶' }, { '7', '⁷' }, { '8', '⁸' }, { '9', '⁹' },
            { '+', '⁺' }, { '-', '⁻' }, { '=', '⁼' }, { '(', '⁽' }, { ')', '⁾' },
            { 'n', 'ⁿ' }, { 'i', 'ⁱ' }, { 'a', 'ᵃ' }, { 'b', 'ᵇ' }, { 'c', 'ᶜ' },
            { 'd', 'ᵈ' }, { 'e', 'ᵉ' }, { 'g', 'ᵍ' }, { 'h', 'ʰ' }, { 'j', 'ʲ' },
            { 'k', 'ᵏ' }, { 'l', 'ˡ' }, { 'm', 'ᵐ' }, { 'o', 'ᵒ' }, { 'p', 'ᵖ' },
            { 'r', 'ʳ' }, { 's', 'ˢ' }, { 't', 'ᵗ' }, { 'u', 'ᵘ' }, { 'v', 'ᵛ' },
            { 'w', 'ʷ' }, { 'x', 'ˣ' }, { 'y', 'ʸ' }, { 'z', 'ᙆ' }
        };

        private static readonly Dictionary<char, char> SubscriptMap = new Dictionary<char, char>
        {
            { '0', '₀' }, { '1', '₁' }, { '2', '₂' }, { '3', '₃' }, { '4', '₄' },
            { '5', '₅' }, { '6', '₆' }, { '7', '₇' }, { '8', '₈' }, { '9', '₉' },
            { '+', '₊' }, { '-', '₋' }, { '=', '₌' }, { '(', '₍' }, { ')', '₎' },
            { 'a', 'ₐ' }, { 'e', 'ₑ' }, { 'h', 'ₕ' }, { 'i', 'ᵢ' }, { 'j', 'ⱼ' },
            { 'k', 'ₖ' }, { 'l', 'ₗ' }, { 'm', 'ₘ' }, { 'n', 'ₙ' }, { 'o', 'ₒ' },
            { 'p', 'ₚ' }, { 'r', 'ᵣ' }, { 's', 'ₛ' }, { 't', 'ₜ' }, { 'u', 'ᵤ' },
            { 'v', 'ᵥ' }, { 'x', 'ₓ' }, { 'y', 'ᵧ' }, { 'z', 'ᵤ' }
        };

        // ========== LaTeX 转换 ==========
        public static string ConvertLatexToUnicode(string latex)
        {
            if (string.IsNullOrEmpty(latex)) return latex;
            latex = latex.Replace("\n", "").Replace("\r", "");
            int i = 0;
            return ParseExpression(latex, ref i);
        }

        private static string ParseExpression(string s, ref int i)
        {
            var sb = new StringBuilder();
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '\\')
                {
                    int start = i;
                    i++;
                    while (i < s.Length && char.IsLetter(s[i])) i++;
                    string cmd = s.Substring(start, i - start);
                    if (cmd == "\\frac")
                    {
                        sb.Append(ParseFraction(s, ref i));
                        continue;
                    }
                    else if (LatexToUnicode.TryGetValue(cmd, out string unicode))
                        sb.Append(unicode);
                    else
                        sb.Append(cmd);
                    continue;
                }
                else if (c == '^')
                {
                    i++;
                    string sup = ParseGroup(s, ref i);
                    sb.Append(ConvertToSuperscript(sup));
                    continue;
                }
                else if (c == '_')
                {
                    i++;
                    string sub = ParseGroup(s, ref i);
                    sb.Append(ConvertToSubscript(sub));
                    continue;
                }
                else if (c == '{')
                {
                    i++;
                    sb.Append(ParseExpression(s, ref i));
                    if (i < s.Length && s[i] == '}') i++;
                    continue;
                }
                else if (c == '}')
                {
                    break;
                }
                else
                {
                    sb.Append(c);
                    i++;
                }
            }
            return sb.ToString();
        }

        private static string ParseFraction(string s, ref int i)
        {
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            string numerator = "";
            if (i < s.Length && s[i] == '{')
            {
                i++;
                numerator = ParseExpression(s, ref i);
                if (i < s.Length && s[i] == '}') i++;
            }
            else if (i < s.Length)
            {
                numerator = s[i].ToString();
                i++;
            }
            else return "";

            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            string denominator = "";
            if (i < s.Length && s[i] == '{')
            {
                i++;
                denominator = ParseExpression(s, ref i);
                if (i < s.Length && s[i] == '}') i++;
            }
            else if (i < s.Length)
            {
                denominator = s[i].ToString();
                i++;
            }
            else return "";

            return $"({numerator})/({denominator})";
        }

        private static string ParseGroup(string s, ref int i)
        {
            if (i < s.Length && s[i] == '{')
            {
                i++;
                string content = ParseExpression(s, ref i);
                if (i < s.Length && s[i] == '}') i++;
                return content;
            }
            else if (i < s.Length)
            {
                char ch = s[i];
                i++;
                return ch.ToString();
            }
            return "";
        }

        private static string ConvertToSuperscript(string str)
        {
            var sb = new StringBuilder();
            foreach (char c in str)
                sb.Append(SuperscriptMap.TryGetValue(c, out char sup) ? sup : c);
            return sb.ToString();
        }

        private static string ConvertToSubscript(string str)
        {
            var sb = new StringBuilder();
            foreach (char c in str)
                sb.Append(SubscriptMap.TryGetValue(c, out char sub) ? sub : c);
            return sb.ToString();
        }

        // ========== 颜色处理 ==========
        private static string PreprocessColorTags(string input)
        {
            string pattern = @"\{color:([#A-Fa-f0-9]{6}|[#A-Fa-f0-9]{3}|[a-z]+)\}(.*?)\{/color\}";
            return Regex.Replace(input, pattern, match =>
            {
                string color = match.Groups[1].Value;
                string inner = match.Groups[2].Value;
                return $"[COLOR={color}]{inner}[/COLOR]";
            });
        }

        private static string PreprocessCustomColorTags(string input)
        {
            string pattern = @"/\(([^_]+)_\{([^}]+)\}/";
            return Regex.Replace(input, pattern, match =>
            {
                string color = match.Groups[1].Value;
                string text = match.Groups[2].Value;
                return $"[COLOR={color}]{text}[/COLOR]";
            });
        }

        private static List<FormattedText> ParseColorTags(string text)
        {
            var result = new List<FormattedText>();
            var regex = new Regex(@"\[COLOR=([#A-Fa-f0-9]{6}|[#A-Fa-f0-9]{3}|[a-z]+)\](.*?)\[/COLOR\]", RegexOptions.Singleline);
            int lastIndex = 0;
            foreach (Match match in regex.Matches(text))
            {
                if (match.Index > lastIndex)
                {
                    string plain = text.Substring(lastIndex, match.Index - lastIndex);
                    if (!string.IsNullOrEmpty(plain))
                        result.Add(new FormattedText { Text = plain });
                }
                string colorCode = match.Groups[1].Value;
                string inner = match.Groups[2].Value;
                Color color = ColorTranslator.FromHtml(colorCode);
                result.Add(new FormattedText { Text = inner, Color = color });
                lastIndex = match.Index + match.Length;
            }
            if (lastIndex < text.Length)
            {
                string remaining = text.Substring(lastIndex);
                if (!string.IsNullOrEmpty(remaining))
                    result.Add(new FormattedText { Text = remaining });
            }
            return result;
        }

        // ========== LaTeX 占位符 ==========
        private static (string processedText, Dictionary<string, string> placeholders) PreprocessLatex(string input)
        {
            var placeholders = new Dictionary<string, string>();
            int blockIndex = 0, inlineIndex = 0;

            string blockPattern = @"\$\$([\s\S]+?)\$\$|\\\[([\s\S]+?)\\\]";
            string processed = Regex.Replace(input, blockPattern, match =>
            {
                string content = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
                string placeholder = $"<LATEX_BLOCK_{blockIndex++}>";
                placeholders[placeholder] = content.Trim();
                return placeholder;
            });

            string inlinePattern = @"\$([^\$]+?)\$|\\\(([^\\]+?)\\\)";
            processed = Regex.Replace(processed, inlinePattern, match =>
            {
                string content = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
                string placeholder = $"<LATEX_INLINE_{inlineIndex++}>";
                placeholders[placeholder] = content.Trim();
                return placeholder;
            });

            return (processed, placeholders);
        }

        private static List<FormattedText> ReplacePlaceholders(List<FormattedText> parts, Dictionary<string, string> placeholders)
        {
            var newParts = new List<FormattedText>();
            foreach (var part in parts)
            {
                if (part.Text == null) continue;
                string text = part.Text;
                bool replaced = false;
                foreach (var kv in placeholders)
                {
                    if (text.Contains(kv.Key))
                    {
                        replaced = true;
                        string[] segments = text.Split(new[] { kv.Key }, StringSplitOptions.None);
                        for (int i = 0; i < segments.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(segments[i]))
                                newParts.Add(new FormattedText { Text = segments[i], Style = part.Style, Color = part.Color, Url = part.Url });
                            if (i < segments.Length - 1)
                            {
                                string formula = ConvertLatexToUnicode(kv.Value);
                                newParts.Add(new FormattedText { Text = formula, Style = FontStyle.Italic, Color = Color.DarkOrange, IsLatex = true });
                            }
                        }
                        break;
                    }
                }
                if (!replaced)
                    newParts.Add(part);
            }
            return newParts;
        }

        // ========== Markdown 解析入口 ==========
        public static List<Paragraph> ParseMarkdown(string markdown)
        {
            markdown = PreprocessCustomColorTags(markdown);
            markdown = PreprocessColorTags(markdown);

            var (processedMarkdown, placeholders) = PreprocessLatex(markdown);

            var paragraphs = new List<Paragraph>();
            if (string.IsNullOrWhiteSpace(processedMarkdown)) return paragraphs;

            var document = Markdown.Parse(processedMarkdown, Pipeline);
            foreach (var block in document)
            {
                if (block is LinkReferenceDefinitionGroup) continue;
                if (block is ThematicBreakBlock) continue;
                if (block is HtmlBlock) continue;

                if (block is HeadingBlock heading)
                    paragraphs.Add(ParseHeading(heading));
                else if (block is ParagraphBlock paragraph)
                    paragraphs.Add(ParseParagraph(paragraph));
                else if (block is ListBlock listBlock)
                    paragraphs.AddRange(ParseList(listBlock));
                else if (block is CodeBlock codeBlock)
                    paragraphs.Add(ParseCodeBlock(codeBlock));
                else if (block is QuoteBlock quoteBlock)
                    paragraphs.Add(ParseQuoteBlock(quoteBlock));
                else if (block is Table table)
                    paragraphs.Add(ParseTable(table));
                else if (block is ThematicBreakBlock)
                {
                    var hr = new Paragraph();
                    hr.FormattedParts.Add(new FormattedText { Text = "————————————" });
                    paragraphs.Add(hr);
                }
            }

            foreach (var para in paragraphs)
                if (para.FormattedParts.Count > 0)
                    para.FormattedParts = ReplacePlaceholders(para.FormattedParts, placeholders);

            // 强制拆分颜色标记（确保所有 [COLOR=...] 都被解析）
            foreach (var para in paragraphs)
            {
                var newParts = new List<FormattedText>();
                foreach (var part in para.FormattedParts)
                {
                    if (part.Text != null && part.Text.Contains("[COLOR="))
                    {
                        newParts.AddRange(ParseColorTags(part.Text));
                    }
                    else
                    {
                        newParts.Add(part);
                    }
                }
                para.FormattedParts = newParts;
            }

            paragraphs.RemoveAll(p => p.FormattedParts.Count == 0 || p.FormattedParts.All(part => string.IsNullOrWhiteSpace(part.Text)));
            return paragraphs;
        }

        // ========== 辅助解析方法 ==========
        private static Paragraph ParseHeading(HeadingBlock heading)
        {
            var p = new Paragraph { Type = ParagraphType.Heading, HeadingLevel = heading.Level };
            p.FormattedParts.AddRange(ParseInline(heading.Inline));
            return p;
        }

        private static Paragraph ParseParagraph(ParagraphBlock paragraph)
        {
            var p = new Paragraph();
            p.FormattedParts.AddRange(ParseInline(paragraph.Inline));
            return p;
        }

        private static IEnumerable<Paragraph> ParseList(ListBlock listBlock)
        {
            var paragraphs = new List<Paragraph>();
            bool isOrdered = listBlock.IsOrdered;
            int itemNumber = 1;
            foreach (var item in listBlock)
            {
                if (item is ListItemBlock listItem)
                {
                    var para = new Paragraph
                    {
                        IsListItem = true,
                        IsOrderedList = isOrdered,
                        ListItemNumber = isOrdered ? itemNumber : 0
                    };
                    foreach (var subBlock in listItem)
                    {
                        if (subBlock is ParagraphBlock subPara)
                        {
                            para.FormattedParts.AddRange(ParseInline(subPara.Inline));
                            paragraphs.Add(para);
                        }
                        else if (subBlock is ListBlock subList)
                        {
                            paragraphs.AddRange(ParseList(subList));
                        }
                    }
                    if (isOrdered) itemNumber++;
                }
            }
            return paragraphs;
        }

        private static Paragraph ParseCodeBlock(CodeBlock codeBlock)
        {
            var p = new Paragraph { Type = ParagraphType.Code };
            var code = codeBlock.Lines.ToString();
            p.FormattedParts.Add(new FormattedText { Text = code, Style = FontStyle.Regular, Color = Color.DarkGray });
            return p;
        }

        private static Paragraph ParseQuoteBlock(QuoteBlock quoteBlock)
        {
            var p = new Paragraph { Type = ParagraphType.Quote };
            foreach (var block in quoteBlock)
            {
                if (block is ParagraphBlock subPara)
                    p.FormattedParts.AddRange(ParseInline(subPara.Inline));
            }
            return p;
        }

        private static Paragraph ParseTable(Table table)
        {
            var p = new Paragraph { Type = ParagraphType.Table };
            var sb = new StringBuilder();
            foreach (var row in table)
            {
                if (row is TableRow tableRow)
                {
                    var cells = new List<string>();
                    foreach (var cell in tableRow)
                    {
                        if (cell is TableCell tableCell)
                        {
                            var inlineText = new List<FormattedText>();
                            foreach (var block in tableCell)
                            {
                                if (block is ParagraphBlock paraBlock && paraBlock.Inline != null)
                                    inlineText.AddRange(ParseInline(paraBlock.Inline));
                            }
                            cells.Add(string.Join("", inlineText.Select(t => t.Text)));
                        }
                    }
                    sb.AppendLine(string.Join("\t", cells));
                }
            }
            p.FormattedParts.Add(new FormattedText { Text = sb.ToString() });
            return p;
        }

        private static IEnumerable<FormattedText> ParseInline(ContainerInline inline)
        {
            var parts = new List<FormattedText>();
            if (inline == null) return parts;

            foreach (var child in inline)
            {
                if (child is LiteralInline literal)
                {
                    parts.AddRange(ParseColorTags(literal.Content.ToString()));
                }
                else if (child is EmphasisInline emphasis)
                {
                    var subText = string.Join("", emphasis.Descendants().OfType<LiteralInline>().Select(l => l.Content.ToString()));
                    FontStyle style = FontStyle.Regular;
                    if (emphasis.DelimiterChar == '*')
                    {
                        if (emphasis.DelimiterCount == 1) style = FontStyle.Italic;
                        else if (emphasis.DelimiterCount == 2) style = FontStyle.Bold;
                        else if (emphasis.DelimiterCount >= 3) style = FontStyle.Bold | FontStyle.Italic;
                    }
                    else if (emphasis.DelimiterChar == '_')
                    {
                        if (emphasis.DelimiterCount == 1) style = FontStyle.Italic;
                        else if (emphasis.DelimiterCount == 2) style = FontStyle.Bold;
                        else if (emphasis.DelimiterCount >= 3) style = FontStyle.Bold | FontStyle.Italic;
                    }
                    parts.Add(new FormattedText { Text = subText, Style = style });
                }
                else if (child is CodeInline code)
                {
                    parts.Add(new FormattedText { Text = code.Content, Style = FontStyle.Regular, Color = Color.DarkRed });
                }
                else if (child is LinkInline link)
                {
                    var linkText = string.Join("", link.Descendants().OfType<LiteralInline>().Select(l => l.Content.ToString()));
                    if (link.IsImage)
                        parts.Add(new FormattedText { Text = $"[图片: {linkText}]", Style = FontStyle.Regular, Color = Color.DarkGreen });
                    else
                        parts.Add(new FormattedText { Text = linkText, Style = FontStyle.Underline, Color = Color.Blue, Url = link.Url });
                }
                else if (child is LineBreakInline)
                {
                    parts.Add(new FormattedText { Text = Environment.NewLine });
                }
                else if (child is HtmlInline)
                {
                    // 忽略 HTML
                }
                else
                {
                    parts.Add(new FormattedText { Text = child.ToString() });
                }
            }
            return parts;
        }
    }
}