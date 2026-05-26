/**
 * Command line tool that formats a cell range (font style, color, background
 * color, etc.). Implements "format_range" from README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormatRangeTool {

    /**
     * Arguments (all values passed as strings):
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>startCell</li>
     *     <li>endCell</li>
     *     <li>bold (true/false)</li>
     *     <li>italic (true/false)</li>
     *     <li>fontSize (integer, 0 to keep current)</li>
     *     <li>fontColor (RGB hex like "#FF0000" or "" to skip)</li>
     *     <li>bgColor (RGB hex like "#FFFF00" or "" to skip)</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 9) {
            System.err.println("Usage: FormatRangeTool <filePath> <sheetName> <startCell> <endCell> <bold> <italic> <fontSize> <fontColor> <bgColor>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String startCell = args[2];
        String endCell = args[3];
        boolean bold = Boolean.parseBoolean(args[4]);
        boolean italic = Boolean.parseBoolean(args[5]);
        int fontSize = Integer.parseInt(args[6]);
        String fontColor = args[7];
        String bgColor = args[8];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            CellRangeAddress range = ExcelRangeUtils.parseRange(startCell + ":" + endCell);

            // POI 4+ requires an IndexedColorMap for reliable RGB color handling in XSSF.
            IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();

            // Parse colors once; they can be reused across styles.
            XSSFColor fontXssfColor = null;
            if (!fontColor.isEmpty()) {
                fontXssfColor = new XSSFColor(parseRgbColor(fontColor), colorMap);
            }

            XSSFColor bgXssfColor = null;
            if (!bgColor.isEmpty()) {
                bgXssfColor = new XSSFColor(parseRgbColor(bgColor), colorMap);
            }

            // IMPORTANT:
            // Do NOT apply a single newly-created style to all cells.
            // That would wipe existing alignment (e.g., TOP -> default BOTTOM), borders, wraps, etc.
            // Instead, clone the existing style per "base style" and apply only requested changes.
            Map<Short, XSSFCellStyle> styleCache = new HashMap<>();

            for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    row = sheet.createRow(r);
                }
                for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                    Cell cell = row.getCell(c);
                    boolean created = false;
                    if (cell == null) {
                        cell = row.createCell(c);
                        created = true;
                    }

                    // If we just created the cell, inherit the best-effort visible style first.
                    // Otherwise, a newly created cell defaults to vertical=BOTTOM etc.
                    if (created && (cell.getCellStyle() == null || cell.getCellStyle().getIndex() == 0)) {
                        applyBestEffortStyle(sheet, r, c, cell);
                    }

                    XSSFCellStyle baseStyle = (XSSFCellStyle) cell.getCellStyle();
                    short baseIdx = baseStyle.getIndex();

                    XSSFCellStyle derived = styleCache.get(baseIdx);
                    if (derived == null) {
                        derived = workbook.createCellStyle();
                        derived.cloneStyleFrom(baseStyle);

                        // Preserve existing font attributes, then override only requested ones.
                        XSSFFont baseFont = workbook.getFontAt(baseStyle.getFontIndexAsInt());
                        XSSFFont newFont = cloneFont(workbook, baseFont);

                        newFont.setBold(bold);
                        newFont.setItalic(italic);
                        if (fontSize > 0) {
                            newFont.setFontHeightInPoints((short) fontSize);
                        }
                        if (fontXssfColor != null) {
                            newFont.setColor(fontXssfColor);
                        }

                        derived.setFont(newFont);

                        if (bgXssfColor != null) {
                            derived.setFillForegroundColor(bgXssfColor);
                            derived.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        }

                        styleCache.put(baseIdx, derived);
                    }

                    cell.setCellStyle(derived);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }

    private static XSSFFont cloneFont(XSSFWorkbook workbook, XSSFFont src) {
        XSSFFont dst = workbook.createFont();
        if (src == null) {
            return dst;
        }

        // Copy typical properties. (POI doesn't provide a direct "clone font" helper.)
        dst.setFontName(src.getFontName());
        dst.setFontHeight(src.getFontHeight());
        dst.setUnderline(src.getUnderline());
        dst.setStrikeout(src.getStrikeout());
        dst.setTypeOffset(src.getTypeOffset());
        dst.setCharSet(src.getCharSet());
        dst.setFamily(src.getFamily());
        dst.setBold(src.getBold());
        dst.setItalic(src.getItalic());

        // Preserve color if present; caller may override later.
        try {
            XSSFColor c = src.getXSSFColor();
            if (c != null) {
                dst.setColor(c);
            } else {
                dst.setColor(src.getColor());
            }
        } catch (Exception ignored) {
            try {
                dst.setColor(src.getColor());
            } catch (Exception ignored2) {
                // ignore
            }
        }

        return dst;
    }

    /**
     * Best-effort style inheritance to keep the "Excel visible look" when new cells are created.
     *
     * Priority:
     *  1) Nearest above non-null cell style (scan upward in the same column)
     *  2) Left cell style (same row, previous column)
     *  3) Column style
     *  4) Row style
     */
    private static void applyBestEffortStyle(Sheet sheet, int rowIndex, int colIndex, Cell target) {
        CellStyle style = null;

        // 1) Scan upward
        for (int r = rowIndex - 1; r >= 0; r--) {
            Row aboveRow = sheet.getRow(r);
            if (aboveRow == null) {
                continue;
            }
            Cell above = aboveRow.getCell(colIndex);
            if (above != null) {
                style = above.getCellStyle();
                break;
            }
        }

        // 2) Left
        if (style == null && colIndex > 0) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell left = row.getCell(colIndex - 1);
                if (left != null) {
                    style = left.getCellStyle();
                }
            }
        }

        // 3) Column style
        if (style == null) {
            try {
                style = sheet.getColumnStyle(colIndex);
            } catch (Exception ignored) {
                // ignore
            }
        }

        // 4) Row style
        if (style == null) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                style = row.getRowStyle();
            }
        }

        if (style != null) {
            target.setCellStyle(style);
        }
    }

    /**
     * Parses a color expressed as "#RRGGBB" into a byte array suitable
     * for {@link XSSFColor}.
     */
    private static byte[] parseRgbColor(String hex) {
        if (hex == null || !hex.matches("#?[0-9A-Fa-f]{6}")) {
            throw new IllegalArgumentException("Invalid RGB color: " + hex);
        }
        String v = hex.startsWith("#") ? hex.substring(1) : hex;
        int r = Integer.parseInt(v.substring(0, 2), 16);
        int g = Integer.parseInt(v.substring(2, 4), 16);
        int b = Integer.parseInt(v.substring(4, 6), 16);
        return new byte[]{(byte) r, (byte) g, (byte) b};
    }
}
