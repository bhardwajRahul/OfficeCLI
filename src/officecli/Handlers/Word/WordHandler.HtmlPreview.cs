// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Generate a self-contained HTML file that previews the Word document
    /// with formatting, tables, images, and lists.
    /// </summary>
    public string ViewAsHtml()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "<html><body><p>(empty document)</p></body></html>";

        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\">");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        sb.AppendLine($"<title>{HtmlEncode(Path.GetFileName(_filePath))}</title>");
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateWordCss());
        sb.AppendLine("</style>");
        // KaTeX for math rendering
        sb.AppendLine("<link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css\">");
        sb.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js\"></script>");
        sb.AppendLine("<script defer src=\"https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/contrib/auto-render.min.js\"></script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        // Page container
        var (pageWidthCm, _, _) = GetPageDimensions();
        var maxW = $"max-width:{pageWidthCm:0.##}cm";

        sb.AppendLine($"<div class=\"page\" style=\"{maxW}\">");

        // Render header
        RenderHeaderFooterHtml(sb, isHeader: true);

        // Render body elements
        RenderBodyHtml(sb, body);

        // Render footer
        RenderHeaderFooterHtml(sb, isHeader: false);

        sb.AppendLine("</div>"); // page

        // KaTeX auto-render script
        sb.AppendLine("<script>");
        sb.AppendLine("document.addEventListener('DOMContentLoaded',function(){");
        sb.AppendLine("  if(typeof renderMathInElement!=='undefined'){");
        sb.AppendLine("    renderMathInElement(document.body,{delimiters:[");
        sb.AppendLine("      {left:'$$',right:'$$',display:true},");
        sb.AppendLine("      {left:'$',right:'$',display:false}");
        sb.AppendLine("    ],throwOnError:false});");
        sb.AppendLine("  }");
        sb.AppendLine("});");
        sb.AppendLine("</script>");

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        return sb.ToString();
    }

    // ==================== Page Dimensions ====================

    private (double widthCm, double marginLeftCm, double marginRightCm) GetPageDimensions()
    {
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
        var pageSize = sectPr?.GetFirstChild<PageSize>();
        var pageMargin = sectPr?.GetFirstChild<PageMargin>();

        // Default A4: 21cm width
        double widthTwips = pageSize?.Width?.Value != null ? (double)pageSize.Width.Value : 11906;
        double marginLeftTwips = pageMargin?.Left?.Value != null ? (double)pageMargin.Left.Value : 1440;
        double marginRightTwips = pageMargin?.Right?.Value != null ? (double)pageMargin.Right.Value : 1440;

        return (
            widthTwips * 2.54 / 1440.0,
            marginLeftTwips * 2.54 / 1440.0,
            marginRightTwips * 2.54 / 1440.0
        );
    }

    // ==================== Header / Footer ====================

    private void RenderHeaderFooterHtml(StringBuilder sb, bool isHeader)
    {
        var cssClass = isHeader ? "doc-header" : "doc-footer";

        if (isHeader)
        {
            var headerParts = _doc.MainDocumentPart?.HeaderParts;
            if (headerParts == null) return;
            foreach (var hp in headerParts)
            {
                var paragraphs = hp.Header?.Elements<Paragraph>().ToList();
                if (paragraphs == null || paragraphs.Count == 0) continue;
                if (paragraphs.All(p => string.IsNullOrWhiteSpace(GetParagraphText(p)))) continue;
                sb.AppendLine($"<div class=\"{cssClass}\">");
                foreach (var para in paragraphs) RenderParagraphHtml(sb, para);
                sb.AppendLine("</div>");
                break;
            }
        }
        else
        {
            var footerParts = _doc.MainDocumentPart?.FooterParts;
            if (footerParts == null) return;
            foreach (var fp in footerParts)
            {
                var paragraphs = fp.Footer?.Elements<Paragraph>().ToList();
                if (paragraphs == null || paragraphs.Count == 0) continue;
                if (paragraphs.All(p => string.IsNullOrWhiteSpace(GetParagraphText(p)))) continue;
                sb.AppendLine($"<div class=\"{cssClass}\">");
                foreach (var para in paragraphs) RenderParagraphHtml(sb, para);
                sb.AppendLine("</div>");
                break;
            }
        }
    }

    // ==================== Body Rendering ====================

    private void RenderBodyHtml(StringBuilder sb, Body body)
    {
        var elements = GetBodyElements(body).ToList();
        // Track list state for proper HTML list rendering
        string? currentListType = null; // "bullet" or "ordered"
        int currentListLevel = 0;
        var listStack = new Stack<string>(); // track nested list tags

        // Detect duplicate content from text boxes (MC AlternateContent)
        // Word stores text box content in both mc:AlternateContent and as flattened runs.
        // We track rendered text to skip duplicates.
        var renderedTextHashes = new HashSet<string>();

        foreach (var element in elements)
        {
            if (element is Paragraph para)
            {
                // Check for display equation
                var oMathPara = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathPara != null)
                {
                    CloseAllLists(sb, listStack, ref currentListType);
                    var latex = FormulaParser.ToLatex(oMathPara);
                    sb.AppendLine($"<div class=\"equation\">$${HtmlEncode(latex)}$$</div>");
                    continue;
                }

                // De-duplicate: if this paragraph's text was already rendered (text box duplicate), skip
                var paraText = GetParagraphText(para).Trim();
                if (!string.IsNullOrEmpty(paraText) && paraText.Length > 10)
                {
                    var hash = paraText;
                    if (renderedTextHashes.Contains(hash))
                        continue; // skip duplicate
                    renderedTextHashes.Add(hash);
                }

                // Check if this is a list item
                var listStyle = GetParagraphListStyle(para);
                if (listStyle != null)
                {
                    var ilvl = para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
                    var tag = listStyle == "bullet" ? "ul" : "ol";

                    // Adjust nesting
                    while (listStack.Count > ilvl + 1)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                    }
                    while (listStack.Count < ilvl + 1)
                    {
                        sb.AppendLine($"<{tag}>");
                        listStack.Push(tag);
                    }
                    // If same level but different list type, swap
                    if (listStack.Count > 0 && listStack.Peek() != tag)
                    {
                        sb.AppendLine($"</{listStack.Pop()}>");
                        sb.AppendLine($"<{tag}>");
                        listStack.Push(tag);
                    }

                    currentListType = listStyle;
                    currentListLevel = ilvl;
                    sb.Append("<li");
                    var paraStyle = GetParagraphInlineCss(para, isListItem: true);
                    if (!string.IsNullOrEmpty(paraStyle))
                        sb.Append($" style=\"{paraStyle}\"");
                    sb.Append(">");
                    RenderParagraphContentHtml(sb, para);
                    sb.AppendLine("</li>");
                    continue;
                }

                // Not a list — close any open lists
                CloseAllLists(sb, listStack, ref currentListType);

                // Check for heading
                var styleName = GetStyleName(para);
                var headingLevel = 0;
                if (styleName.Contains("Heading") || styleName.Contains("标题")
                    || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase))
                {
                    headingLevel = GetHeadingLevel(styleName);
                    if (headingLevel < 1) headingLevel = 1;
                    if (headingLevel > 6) headingLevel = 6;
                }
                else if (styleName == "Title")
                    headingLevel = 1;
                else if (styleName == "Subtitle")
                    headingLevel = 2;

                if (headingLevel > 0)
                {
                    sb.Append($"<h{headingLevel}");
                    var hStyle = GetParagraphInlineCss(para);
                    if (!string.IsNullOrEmpty(hStyle))
                        sb.Append($" style=\"{hStyle}\"");
                    sb.Append(">");
                    RenderParagraphContentHtml(sb, para);
                    sb.AppendLine($"</h{headingLevel}>");
                }
                else
                {
                    // Normal paragraph
                    var text = GetParagraphText(para);
                    var runs = GetAllRuns(para);
                    var mathElements = FindMathElements(para);

                    // Empty paragraph = spacing break
                    if (runs.Count == 0 && mathElements.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        sb.AppendLine("<p class=\"empty\">&nbsp;</p>");
                        continue;
                    }

                    // Inline equation only
                    if (mathElements.Count > 0 && runs.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        var latex = string.Concat(mathElements.Select(FormulaParser.ToLatex));
                        sb.AppendLine($"<div class=\"equation\">$${HtmlEncode(latex)}$$</div>");
                        continue;
                    }

                    sb.Append("<p");
                    var pStyle = GetParagraphInlineCss(para);
                    if (!string.IsNullOrEmpty(pStyle))
                        sb.Append($" style=\"{pStyle}\"");
                    sb.Append(">");
                    RenderParagraphContentHtml(sb, para);
                    sb.AppendLine("</p>");
                }
            }
            else if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                CloseAllLists(sb, listStack, ref currentListType);
                var latex = FormulaParser.ToLatex(element);
                sb.AppendLine($"<div class=\"equation\">$${HtmlEncode(latex)}$$</div>");
            }
            else if (element is Table table)
            {
                CloseAllLists(sb, listStack, ref currentListType);
                RenderTableHtml(sb, table);
            }
            else if (element is SectionProperties)
            {
                // Skip — section properties are not visual content
            }
        }

        CloseAllLists(sb, listStack, ref currentListType);
    }

    private static void CloseAllLists(StringBuilder sb, Stack<string> listStack, ref string? currentListType)
    {
        while (listStack.Count > 0)
            sb.AppendLine($"</{listStack.Pop()}>");
        currentListType = null;
    }

    // ==================== Paragraph Content ====================

    private void RenderParagraphHtml(StringBuilder sb, Paragraph para)
    {
        sb.Append("<p");
        var pStyle = GetParagraphInlineCss(para);
        if (!string.IsNullOrEmpty(pStyle))
            sb.Append($" style=\"{pStyle}\"");
        sb.Append(">");
        RenderParagraphContentHtml(sb, para);
        sb.AppendLine("</p>");
    }

    private void RenderParagraphContentHtml(StringBuilder sb, Paragraph para)
    {
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
            {
                RenderRunHtml(sb, run, para);
            }
            else if (child is Hyperlink hyperlink)
            {
                var relId = hyperlink.Id?.Value;
                string? url = null;
                if (relId != null)
                {
                    try
                    {
                        url = _doc.MainDocumentPart?.HyperlinkRelationships
                            .FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                    }
                    catch { }
                    if (url == null)
                    {
                        try
                        {
                            url = _doc.MainDocumentPart?.ExternalRelationships
                                .FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                        }
                        catch { }
                    }
                }

                if (url != null)
                    sb.Append($"<a href=\"{HtmlEncode(url)}\" target=\"_blank\">");

                foreach (var hRun in hyperlink.Elements<Run>())
                    RenderRunHtml(sb, hRun, para);

                if (url != null)
                    sb.Append("</a>");
            }
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
            {
                var latex = FormulaParser.ToLatex(child);
                sb.Append($"${HtmlEncode(latex)}$");
            }
        }
    }

    // ==================== Run Rendering ====================

    private void RenderRunHtml(StringBuilder sb, Run run, Paragraph para)
    {
        // Check for image
        var drawing = run.GetFirstChild<Drawing>();
        if (drawing != null)
        {
            RenderDrawingHtml(sb, drawing);
            return;
        }

        // Check for break
        var br = run.GetFirstChild<Break>();
        if (br != null)
        {
            if (br.Type?.Value == BreakValues.Page)
                sb.Append("<hr class=\"page-break\">");
            else
                sb.Append("<br>");
        }

        // Check for tab
        var tab = run.GetFirstChild<TabChar>();

        var text = GetRunText(run);
        if (string.IsNullOrEmpty(text) && tab == null) return;

        var rProps = ResolveEffectiveRunProperties(run, para);
        var style = GetRunInlineCss(rProps);

        var needsSpan = !string.IsNullOrEmpty(style);
        if (needsSpan)
            sb.Append($"<span style=\"{style}\">");

        if (tab != null)
            sb.Append("&emsp;");

        sb.Append(HtmlEncode(text));

        if (needsSpan)
            sb.Append("</span>");
    }

    // ==================== Image Rendering ====================

    private void RenderDrawingHtml(StringBuilder sb, Drawing drawing)
    {
        // Try to find the blip (embedded image reference)
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value == null) return;

        var relId = blip.Embed.Value;
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return;

        try
        {
            var imagePart = mainPart.GetPartById(relId) as ImagePart;
            if (imagePart == null) return;

            var contentType = imagePart.ContentType;
            using var stream = imagePart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var base64 = Convert.ToBase64String(ms.ToArray());

            // Get dimensions
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault()
                ?? drawing.Descendants<A.Extents>().FirstOrDefault() as OpenXmlElement;
            string widthAttr = "", heightAttr = "";
            if (extent is DW.Extent dwExt)
            {
                if (dwExt.Cx?.Value > 0) widthAttr = $" width=\"{dwExt.Cx.Value / 9525}\"";
                if (dwExt.Cy?.Value > 0) heightAttr = $" height=\"{dwExt.Cy.Value / 9525}\"";
            }
            else if (extent is A.Extents aExt)
            {
                if (aExt.Cx?.Value > 0) widthAttr = $" width=\"{aExt.Cx.Value / 9525}\"";
                if (aExt.Cy?.Value > 0) heightAttr = $" height=\"{aExt.Cy.Value / 9525}\"";
            }

            // Get alt text
            var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
            var alt = docProps?.Description?.Value ?? docProps?.Name?.Value ?? "image";

            sb.Append($"<img src=\"data:{contentType};base64,{base64}\" alt=\"{HtmlEncode(alt)}\"{widthAttr}{heightAttr} style=\"max-width:100%;height:auto\">");
        }
        catch
        {
            sb.Append("<span class=\"img-error\">[Image]</span>");
        }
    }

    // ==================== Table Rendering ====================

    private void RenderTableHtml(StringBuilder sb, Table table)
    {
        // Check table-level borders to determine if this is a borderless layout table
        var tblBorders = table.GetFirstChild<TableProperties>()?.TableBorders;
        bool tableBordersNone = IsTableBorderless(tblBorders);

        var tableClass = tableBordersNone ? "borderless" : "";
        sb.AppendLine(string.IsNullOrEmpty(tableClass) ? "<table>" : $"<table class=\"{tableClass}\">");

        // Get column widths from grid
        var tblGrid = table.GetFirstChild<TableGrid>();
        if (tblGrid != null)
        {
            sb.Append("<colgroup>");
            foreach (var col in tblGrid.Elements<GridColumn>())
            {
                var w = col.Width?.Value;
                if (w != null)
                {
                    var px = (int)(double.Parse(w) / 1440.0 * 96); // twips to px
                    sb.Append($"<col style=\"width:{px}px\">");
                }
                else
                {
                    sb.Append("<col>");
                }
            }
            sb.AppendLine("</colgroup>");
        }

        foreach (var row in table.Elements<TableRow>())
        {
            var isHeader = row.TableRowProperties?.GetFirstChild<TableHeader>() != null;
            sb.AppendLine(isHeader ? "<tr class=\"header-row\">" : "<tr>");

            foreach (var cell in row.Elements<TableCell>())
            {
                var tag = isHeader ? "th" : "td";
                var cellStyle = GetTableCellInlineCss(cell, tableBordersNone);

                // Merge attributes
                var attrs = new StringBuilder();
                var gridSpan = cell.TableCellProperties?.GridSpan?.Val?.Value;
                if (gridSpan > 1) attrs.Append($" colspan=\"{gridSpan}\"");

                var vMerge = cell.TableCellProperties?.VerticalMerge;
                if (vMerge != null && vMerge.Val?.Value == MergedCellValues.Restart)
                {
                    // Count rowspan
                    var rowspan = CountRowSpan(table, row, cell);
                    if (rowspan > 1) attrs.Append($" rowspan=\"{rowspan}\"");
                }
                else if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
                {
                    continue; // Skip merged continuation cells
                }

                if (!string.IsNullOrEmpty(cellStyle))
                    attrs.Append($" style=\"{cellStyle}\"");

                sb.Append($"<{tag}{attrs}>");

                // Render cell content — use paragraph tags for multi-paragraph cells
                var cellParagraphs = cell.Elements<Paragraph>().ToList();
                for (int pi = 0; pi < cellParagraphs.Count; pi++)
                {
                    var cellPara = cellParagraphs[pi];
                    var text = GetParagraphText(cellPara);
                    var runs = GetAllRuns(cellPara);

                    if (runs.Count == 0 && string.IsNullOrWhiteSpace(text))
                    {
                        // empty cell paragraph — skip but preserve spacing between paragraphs
                        if (pi > 0 && pi < cellParagraphs.Count - 1)
                            sb.Append("<br>");
                    }
                    else
                    {
                        var pCss = GetParagraphInlineCss(cellPara);
                        if (!string.IsNullOrEmpty(pCss))
                            sb.Append($"<div style=\"{pCss}\">");
                        RenderParagraphContentHtml(sb, cellPara);
                        if (!string.IsNullOrEmpty(pCss))
                            sb.Append("</div>");
                        else if (pi < cellParagraphs.Count - 1)
                            sb.Append("<br>");
                    }
                }

                // Render nested tables
                foreach (var nestedTable in cell.Elements<Table>())
                    RenderTableHtml(sb, nestedTable);

                sb.AppendLine($"</{tag}>");
            }

            sb.AppendLine("</tr>");
        }

        sb.AppendLine("</table>");
    }

    private static bool IsTableBorderless(TableBorders? borders)
    {
        if (borders == null) return false;
        // Check if all borders are none/nil
        return IsBorderNone(borders.TopBorder)
            && IsBorderNone(borders.BottomBorder)
            && IsBorderNone(borders.LeftBorder)
            && IsBorderNone(borders.RightBorder)
            && IsBorderNone(borders.InsideHorizontalBorder)
            && IsBorderNone(borders.InsideVerticalBorder);
    }

    private static bool IsBorderNone(OpenXmlElement? border)
    {
        if (border == null) return true;
        var val = border.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return val is null or "nil" or "none";
    }

    private static int CountRowSpan(Table table, TableRow startRow, TableCell startCell)
    {
        var rows = table.Elements<TableRow>().ToList();
        var startRowIdx = rows.IndexOf(startRow);
        var cellIdx = startRow.Elements<TableCell>().ToList().IndexOf(startCell);
        if (startRowIdx < 0 || cellIdx < 0) return 1;

        int span = 1;
        for (int i = startRowIdx + 1; i < rows.Count; i++)
        {
            var cells = rows[i].Elements<TableCell>().ToList();
            if (cellIdx >= cells.Count) break;

            var vm = cells[cellIdx].TableCellProperties?.VerticalMerge;
            if (vm != null && (vm.Val == null || vm.Val.Value == MergedCellValues.Continue))
                span++;
            else
                break;
        }
        return span;
    }

    // ==================== Inline CSS ====================

    private string GetParagraphInlineCss(Paragraph para, bool isListItem = false)
    {
        var parts = new List<string>();

        var pProps = para.ParagraphProperties;
        if (pProps == null) return ResolveParagraphStyleCss(para);

        // Alignment
        var jc = pProps.Justification?.Val;
        if (jc != null)
        {
            var align = jc.InnerText switch
            {
                "center" => "center",
                "right" or "end" => "right",
                "both" or "distribute" => "justify",
                _ => (string?)null
            };
            if (align != null) parts.Add($"text-align:{align}");
        }

        // Indentation (skip for list items — handled by list nesting)
        if (!isListItem)
        {
            var indent = pProps.Indentation;
            if (indent != null)
            {
                if (indent.Left?.Value is string leftTwips && leftTwips != "0")
                    parts.Add($"margin-left:{TwipsToPx(leftTwips)}px");
                if (indent.Right?.Value is string rightTwips && rightTwips != "0")
                    parts.Add($"margin-right:{TwipsToPx(rightTwips)}px");
                if (indent.FirstLine?.Value is string firstLineTwips && firstLineTwips != "0")
                    parts.Add($"text-indent:{TwipsToPx(firstLineTwips)}px");
                if (indent.Hanging?.Value is string hangTwips && hangTwips != "0")
                    parts.Add($"text-indent:-{TwipsToPx(hangTwips)}px");
            }
        }

        // Spacing
        var spacing = pProps.SpacingBetweenLines;
        if (spacing != null)
        {
            if (spacing.Before?.Value is string beforeTwips && beforeTwips != "0")
                parts.Add($"margin-top:{TwipsToPx(beforeTwips)}px");
            if (spacing.After?.Value is string afterTwips && afterTwips != "0")
                parts.Add($"margin-bottom:{TwipsToPx(afterTwips)}px");
            if (spacing.Line?.Value is string lineVal)
            {
                var rule = spacing.LineRule?.InnerText;
                if (rule == "auto" || rule == null)
                {
                    // Multiplier: value/240 = line spacing ratio
                    if (int.TryParse(lineVal, out var lv))
                        parts.Add($"line-height:{lv / 240.0:0.##}");
                }
                else if (rule == "exact" || rule == "atLeast")
                {
                    parts.Add($"line-height:{TwipsToPx(lineVal)}px");
                }
            }
        }

        // Shading / background (direct or from style)
        var shading = pProps.Shading;
        if (shading?.Fill?.Value is string fill && fill != "auto")
            parts.Add($"background-color:#{fill}");
        else
        {
            // Try to resolve from paragraph style
            var bgFromStyle = ResolveParagraphShadingFromStyle(para);
            if (bgFromStyle != null) parts.Add($"background-color:#{bgFromStyle}");
        }

        // Borders
        var pBdr = pProps.ParagraphBorders;
        if (pBdr != null)
        {
            RenderBorderCss(parts, pBdr.TopBorder, "border-top");
            RenderBorderCss(parts, pBdr.BottomBorder, "border-bottom");
            RenderBorderCss(parts, pBdr.LeftBorder, "border-left");
            RenderBorderCss(parts, pBdr.RightBorder, "border-right");
        }

        return string.Join(";", parts);
    }

    /// <summary>
    /// Resolve paragraph background shading from the style chain.
    /// </summary>
    private string? ResolveParagraphShadingFromStyle(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return null;

        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;

            var shading = style.StyleParagraphProperties?.Shading;
            if (shading?.Fill?.Value is string fill && fill != "auto")
                return fill;

            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    /// <summary>
    /// Resolve paragraph CSS from style chain when no direct paragraph properties.
    /// </summary>
    private string ResolveParagraphStyleCss(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return "";

        var parts = new List<string>();
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;

            var pPr = style.StyleParagraphProperties;
            if (pPr != null)
            {
                var jc = pPr.Justification?.Val;
                if (jc != null && !parts.Any(p => p.StartsWith("text-align")))
                {
                    var align = jc.InnerText switch { "center" => "center", "right" or "end" => "right", "both" => "justify", _ => (string?)null };
                    if (align != null) parts.Add($"text-align:{align}");
                }

                var spacing = pPr.SpacingBetweenLines;
                if (spacing != null)
                {
                    if (spacing.Before?.Value is string b && b != "0" && !parts.Any(p => p.StartsWith("margin-top")))
                        parts.Add($"margin-top:{TwipsToPx(b)}px");
                    if (spacing.After?.Value is string a && a != "0" && !parts.Any(p => p.StartsWith("margin-bottom")))
                        parts.Add($"margin-bottom:{TwipsToPx(a)}px");
                    if (spacing.Line?.Value is string lv && !parts.Any(p => p.StartsWith("line-height")))
                    {
                        var rule = spacing.LineRule?.InnerText;
                        if ((rule == "auto" || rule == null) && int.TryParse(lv, out var val))
                            parts.Add($"line-height:{val / 240.0:0.##}");
                    }
                }

                var shading = pPr.Shading;
                if (shading?.Fill?.Value is string fill && fill != "auto" && !parts.Any(p => p.StartsWith("background")))
                    parts.Add($"background-color:#{fill}");
            }

            currentStyleId = style.BasedOn?.Val?.Value;
        }
        return string.Join(";", parts);
    }

    private static string GetRunInlineCss(RunProperties? rProps)
    {
        if (rProps == null) return "";
        var parts = new List<string>();

        // Font
        var fonts = rProps.RunFonts;
        var font = fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
        if (font != null) parts.Add($"font-family:'{CssSanitize(font)}'");

        // Size (stored as half-points)
        var size = rProps.FontSize?.Val?.Value;
        if (size != null && int.TryParse(size, out var halfPts))
            parts.Add($"font-size:{halfPts / 2.0:0.##}pt");

        // Bold
        if (rProps.Bold != null)
            parts.Add("font-weight:bold");

        // Italic
        if (rProps.Italic != null)
            parts.Add("font-style:italic");

        // Underline
        if (rProps.Underline?.Val != null)
        {
            var ulVal = rProps.Underline.Val.InnerText;
            if (ulVal != "none")
                parts.Add("text-decoration:underline");
        }

        // Strikethrough
        if (rProps.Strike != null)
        {
            var existing = parts.FirstOrDefault(p => p.StartsWith("text-decoration:"));
            if (existing != null)
            {
                parts.Remove(existing);
                parts.Add(existing + " line-through");
            }
            else
            {
                parts.Add("text-decoration:line-through");
            }
        }

        // Color
        var color = rProps.Color?.Val?.Value;
        if (color != null && color != "auto")
            parts.Add($"color:#{color}");

        // Highlight
        var highlight = rProps.Highlight?.Val?.InnerText;
        if (highlight != null)
        {
            var hlColor = HighlightToCssColor(highlight);
            if (hlColor != null) parts.Add($"background-color:{hlColor}");
        }

        // Superscript / Subscript
        var vertAlign = rProps.VerticalTextAlignment?.Val;
        if (vertAlign != null)
        {
            if (vertAlign.InnerText == "superscript")
                parts.Add("vertical-align:super;font-size:smaller");
            else if (vertAlign.InnerText == "subscript")
                parts.Add("vertical-align:sub;font-size:smaller");
        }

        return string.Join(";", parts);
    }

    private string GetTableCellInlineCss(TableCell cell, bool tableBordersNone)
    {
        var parts = new List<string>();
        var tcPr = cell.TableCellProperties;

        // If table-level borders are none, explicitly set border:none on cells
        if (tableBordersNone)
            parts.Add("border:none");

        if (tcPr == null) return string.Join(";", parts);

        // Shading / fill
        var shading = tcPr.Shading;
        if (shading?.Fill?.Value is string fill && fill != "auto")
            parts.Add($"background-color:#{fill}");

        // Vertical alignment
        var vAlign = tcPr.TableCellVerticalAlignment?.Val;
        if (vAlign != null)
        {
            var va = vAlign.InnerText switch
            {
                "center" => "middle",
                "bottom" => "bottom",
                _ => (string?)null
            };
            if (va != null) parts.Add($"vertical-align:{va}");
        }

        // Cell borders (override table-level setting if cell has its own)
        var tcBorders = tcPr.TableCellBorders;
        if (tcBorders != null)
        {
            // Remove the table-level border:none if cell has specific borders
            if (tableBordersNone)
                parts.Remove("border:none");
            RenderBorderCss(parts, tcBorders.TopBorder, "border-top");
            RenderBorderCss(parts, tcBorders.BottomBorder, "border-bottom");
            RenderBorderCss(parts, tcBorders.LeftBorder, "border-left");
            RenderBorderCss(parts, tcBorders.RightBorder, "border-right");
        }

        // Cell width
        var width = tcPr.TableCellWidth?.Width?.Value;
        if (width != null && int.TryParse(width, out var w))
        {
            var type = tcPr.TableCellWidth?.Type?.InnerText;
            if (type == "dxa")
                parts.Add($"width:{w / 1440.0 * 96:0}px");
            else if (type == "pct")
                parts.Add($"width:{w / 50.0:0.#}%");
        }

        // Padding
        var margins = tcPr.TableCellMargin;
        if (margins != null)
        {
            var padTop = margins.TopMargin?.Width?.Value;
            var padBot = margins.BottomMargin?.Width?.Value;
            var padLeft = margins.LeftMargin?.Width?.Value ?? margins.StartMargin?.Width?.Value;
            var padRight = margins.RightMargin?.Width?.Value ?? margins.EndMargin?.Width?.Value;
            if (padTop != null || padBot != null || padLeft != null || padRight != null)
            {
                parts.Add($"padding:{TwipsToPxStr(padTop ?? "0")} {TwipsToPxStr(padRight ?? "0")} {TwipsToPxStr(padBot ?? "0")} {TwipsToPxStr(padLeft ?? "0")}");
            }
        }

        return string.Join(";", parts);
    }

    // ==================== CSS Helpers ====================

    private static void RenderBorderCss(List<string> parts, OpenXmlElement? border, string cssProp)
    {
        if (border == null) return;
        var val = border.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (val == null || val == "nil" || val == "none") return;

        var sz = border.GetAttributes().FirstOrDefault(a => a.LocalName == "sz").Value;
        var color = border.GetAttributes().FirstOrDefault(a => a.LocalName == "color").Value;

        var width = sz != null && int.TryParse(sz, out var s) ? $"{Math.Max(1, s / 8.0):0.#}px" : "1px";
        var style = val switch
        {
            "single" => "solid",
            "double" => "double",
            "dashed" or "dashSmallGap" => "dashed",
            "dotted" => "dotted",
            _ => "solid"
        };
        var cssColor = (color != null && color != "auto") ? $"#{color}" : "#000";

        parts.Add($"{cssProp}:{width} {style} {cssColor}");
    }

    private static int TwipsToPx(string twipsStr)
    {
        if (!int.TryParse(twipsStr, out var twips)) return 0;
        return (int)(twips / 1440.0 * 96);
    }

    private static string TwipsToPxStr(string twipsStr)
    {
        return $"{TwipsToPx(twipsStr)}px";
    }

    private static string? HighlightToCssColor(string highlight) => highlight.ToLowerInvariant() switch
    {
        "yellow" => "#FFFF00",
        "green" => "#00FF00",
        "cyan" => "#00FFFF",
        "magenta" => "#FF00FF",
        "blue" => "#0000FF",
        "red" => "#FF0000",
        "darkblue" => "#00008B",
        "darkcyan" => "#008B8B",
        "darkgreen" => "#006400",
        "darkmagenta" => "#8B008B",
        "darkred" => "#8B0000",
        "darkyellow" => "#808000",
        "darkgray" => "#A9A9A9",
        "lightgray" => "#D3D3D3",
        "black" => "#000000",
        "white" => "#FFFFFF",
        _ => null
    };

    private static string CssSanitize(string value) =>
        Regex.Replace(value, @"[""'\\<>&;{}]", "");

    private static string HtmlEncode(string? text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }

    // ==================== CSS Stylesheet ====================

    private static string GenerateWordCss() => """
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            background: #f0f0f0;
            font-family: 'Microsoft YaHei', 'Segoe UI', -apple-system, 'PingFang SC', 'Hiragino Sans GB', sans-serif;
            color: #333;
            padding: 20px;
        }
        .page {
            background: white;
            margin: 0 auto 40px;
            padding: 2.54cm 0 2.54cm 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            border-radius: 4px;
            min-height: 29.7cm;
            line-height: 1.5;
            font-size: 10.5pt;
        }
        .doc-header, .doc-footer {
            padding: 0 2.54cm;
            color: #888;
            font-size: 9pt;
            border-bottom: 1px solid #e0e0e0;
            margin-bottom: 1em;
            padding-bottom: 0.5em;
        }
        .doc-footer {
            border-bottom: none;
            border-top: 1px solid #e0e0e0;
            margin-top: 1em;
            padding-top: 0.5em;
            margin-bottom: 0;
        }
        h1, h2, h3, h4, h5, h6 {
            padding: 0.3em 2.54cm;
            line-height: 1.4;
        }
        h1 { font-size: 22pt; margin-top: 0.5em; margin-bottom: 0.3em; }
        h2 { font-size: 16pt; margin-top: 0.4em; margin-bottom: 0.2em; }
        h3 { font-size: 13pt; margin-top: 0.3em; margin-bottom: 0.2em; }
        h4 { font-size: 11pt; margin-top: 0.2em; margin-bottom: 0.1em; }
        h5 { font-size: 10pt; }
        h6 { font-size: 9pt; }
        p {
            padding: 0 2.54cm;
            margin: 0.1em 0;
        }
        p.empty {
            margin: 0;
            padding: 0 2.54cm;
            line-height: 0.8;
            font-size: 6pt;
        }
        a { color: #2B579A; }
        a:hover { color: #1a3c6e; }
        ul, ol {
            padding-left: 2em;
            margin: 0.2em 0 0.2em 2.54cm;
        }
        li {
            margin: 0.1em 0;
        }
        .equation {
            text-align: center;
            padding: 0.5em 2.54cm;
            overflow-x: auto;
        }
        img {
            max-width: 100%;
            height: auto;
        }
        .img-error {
            color: #999;
            font-style: italic;
        }
        table {
            border-collapse: collapse;
            margin: 0.3em 2.54cm;
            font-size: 10.5pt;
            width: calc(100% - 5.08cm);
        }
        table.borderless {
            border: none;
        }
        table.borderless td, table.borderless th {
            border: none;
            padding: 2px 6px;
        }
        th, td {
            border: 1px solid #bbb;
            padding: 4px 8px;
            text-align: left;
            vertical-align: top;
        }
        th {
            background: #f0f0f0;
            font-weight: 600;
        }
        .header-row td, .header-row th {
            background: #f0f0f0;
            font-weight: 600;
        }
        hr.page-break {
            border: none;
            border-top: 2px dashed #ccc;
            margin: 2em 2.54cm;
        }
        @media print {
            body { background: white; padding: 0; }
            .page { box-shadow: none; margin: 0; max-width: none; }
            hr.page-break { page-break-after: always; border: none; margin: 0; }
        }
        """;
}
