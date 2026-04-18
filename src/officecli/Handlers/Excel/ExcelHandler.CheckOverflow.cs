// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // CONSISTENCY(text-overflow-check): mirrors PowerPointHandler.CheckShapeTextOverflow.
    // Narrow scope vs PPT: only flags wrapText cells where row height is fixed too small
    // (merged cells, or non-merged cells with explicit customHeight). Skips overflow-right
    // on non-wrapText cells — that is Excel's normal rendering, not a bug.

    /// <summary>
    /// Scan every sheet for cells whose wrapped text cannot fit inside the visible
    /// row-height budget. Returns (path, message) pairs suitable for the `check`
    /// command output. Mirrors PowerPointHandler's CheckShapeTextOverflow pattern.
    /// </summary>
    public List<(string Path, string Message)> CheckAllCellOverflow()
    {
        var issues = new List<(string, string)>();
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;

        foreach (var (sheetName, part) in GetWorksheets(_doc))
        {
            var ws = part.Worksheet;
            if (ws == null) continue;
            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            var mergeMap = BuildMergeMap(ws);
            var colWidths = GetColumnWidths(ws);

            var rowHeights = new Dictionary<int, (double Height, bool Custom)>();
            foreach (var row in sheetData.Elements<Row>())
            {
                int rIdx = (int)(row.RowIndex?.Value ?? 0);
                if (rIdx == 0 || row.Height?.Value == null) continue;
                rowHeights[rIdx] = (row.Height.Value, row.CustomHeight?.Value == true);
            }

            var sheetFmtPr = ws.GetFirstChild<SheetFormatProperties>();
            double defaultRowHeightPt = sheetFmtPr?.DefaultRowHeight?.Value ?? 15.0;
            double defaultColWidthPt = sheetFmtPr?.DefaultColumnWidth?.Value != null
                ? sheetFmtPr.DefaultColumnWidth.Value * 7.0017 * 0.75
                : 8.43 * 7.0017 * 0.75;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    bool isMerged = mergeMap.TryGetValue(cellRef, out var mInfo);
                    if (isMerged && !mInfo.IsAnchor) continue;

                    if (!TryGetCellAlignmentAndFont(cell, stylesheet, out var wrapText, out var fontSizePt))
                        continue;
                    if (!wrapText) continue;

                    var text = GetCellDisplayValue(cell);
                    if (string.IsNullOrEmpty(text)) continue;

                    var (startCol, startRow) = ParseCellReference(cellRef);
                    int startColIdx = ColumnNameToIndex(startCol);
                    int rowSpan = isMerged ? mInfo.RowSpan : 1;
                    int colSpan = isMerged ? mInfo.ColSpan : 1;

                    // Non-merged cells with wrapText default to auto-fit — only flag when
                    // someone explicitly pinned the row height (customHeight="1").
                    if (!isMerged)
                    {
                        if (!rowHeights.TryGetValue(startRow, out var rh) || !rh.Custom)
                            continue;
                    }

                    double usableWidth = 0;
                    for (int c = startColIdx; c < startColIdx + colSpan; c++)
                        usableWidth += colWidths.TryGetValue(c, out var w) ? w : defaultColWidthPt;
                    usableWidth -= 6; // ~3pt side padding total

                    double usableHeight = 0;
                    for (int r = startRow; r < startRow + rowSpan; r++)
                        usableHeight += rowHeights.TryGetValue(r, out var rh2) ? rh2.Height : defaultRowHeightPt;
                    usableHeight -= 4; // ~2pt top/bottom padding total

                    if (usableWidth <= 0 || usableHeight <= 0) continue;

                    double lineHeight = fontSizePt * 1.2;
                    int totalLines = CountWrappedLines(text, fontSizePt, usableWidth);
                    double needed = totalLines * lineHeight;
                    // Require at least ~30% of one line to be clipped. 1-2pt differences
                    // are rendering-metric noise (Excel uses slightly less than 1.2× line height)
                    // and would drown real issues in false positives.
                    if (needed - usableHeight < lineHeight * 0.3) continue;

                    string path = $"/{sheetName}/{cellRef}";
                    string mergeNote = isMerged
                        ? $" (merged {cellRef}:{IndexToColumnName(startColIdx + colSpan - 1)}{startRow + rowSpan - 1})"
                        : "";
                    string suggest;
                    if (isMerged)
                    {
                        double perRowPt = Math.Ceiling((needed + 4) / rowSpan / 5.0) * 5.0;
                        suggest = $"suggest.rowHeight={perRowPt:F0}pt per row (Excel does not auto-fit merged rows)";
                    }
                    else
                    {
                        suggest = "suggest: clear customHeight to let Excel auto-fit";
                    }
                    issues.Add((path,
                        $"text overflow{mergeNote}: {totalLines} lines at {fontSizePt:F1}pt need {needed:F0}pt, usable {usableHeight:F0}pt. {suggest}"));
                }
            }
        }
        return issues;
    }

    private static int CountWrappedLines(string text, double fontSizePt, double usableWidthPt)
    {
        // Newline handling mirrors PowerPointHandler.CheckTextOverflow: both literal
        // and escaped "\n" split into separate paragraphs.
        var paragraphs = text.Replace("\\n", "\n").Split('\n');
        int total = 0;
        foreach (var segment in paragraphs)
        {
            if (segment.Length == 0) { total++; continue; }
            int lines = 1;
            double w = 0;
            foreach (char ch in segment)
            {
                double cw = ParseHelpers.IsCjkOrFullWidth(ch) ? fontSizePt : fontSizePt * 0.55;
                if (w + cw > usableWidthPt && w > 0)
                {
                    lines++;
                    w = cw;
                }
                else
                {
                    w += cw;
                }
            }
            total += lines;
        }
        return total;
    }

    private static bool TryGetCellAlignmentAndFont(
        Cell cell, Stylesheet? stylesheet, out bool wrapText, out double fontSizePt)
    {
        wrapText = false;
        fontSizePt = 11.0; // Excel default body font
        if (stylesheet == null) return true;

        var styleIndex = (int)(cell.StyleIndex?.Value ?? 0);
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null) return true;
        var xfList = cellFormats.Elements<CellFormat>().ToList();
        if (styleIndex >= xfList.Count) return true;
        var xf = xfList[styleIndex];

        wrapText = xf.Alignment?.WrapText?.Value == true;

        var fonts = stylesheet.Fonts;
        if (fonts != null)
        {
            var fontId = (int)(xf.FontId?.Value ?? 0);
            var fontList = fonts.Elements<Font>().ToList();
            if (fontId < fontList.Count)
            {
                var sz = fontList[fontId].FontSize?.Val?.Value;
                if (sz.HasValue) fontSizePt = sz.Value;
            }
        }
        return true;
    }
}
