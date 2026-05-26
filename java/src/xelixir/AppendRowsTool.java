/**
 * Command line tool that appends rows (2D array) to a worksheet at the first
 * empty row, determined by scanning an "anchor" column from top to bottom.
 *
 * Arguments:
 * <ol>
 *     <li>filePath - path to the .xlsx file</li>
 *     <li>sheetName - name of the worksheet</li>
 *     <li>anchorColumn - column letter used to find the first empty row (e.g. "A")</li>
 *     <li>jsonRows - JSON array of arrays representing rows and cells</li>
 * </ol>
 *
 * Notes:
 * - "First empty row" means: the row is missing, or the anchor cell is blank/empty.
 * - When a string cell value starts with "=", it will be written as a formula.
 * - Null values are written as blank cells.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AppendRowsTool {

    public static void main(String[] args) throws Exception {
        if (args.length != 4) {
            System.err.println("Usage: AppendRowsTool <filePath> <sheetName> <anchorColumn> <jsonRows>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String anchorColumn = args[2];
        String jsonRows = args[3];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        int anchorColIndex = CellReference.convertColStringToIndex(anchorColumn);

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = ExcelUtils.getOrCreateSheet(workbook, sheetName);

            JsonArray rows = JsonParser.parseString(jsonRows).getAsJsonArray();

            int startRow = findFirstEmptyRowByAnchorColumn(sheet, anchorColIndex);
            int startCol = 0; // append starts from column A

            Row templateRow = startRow > 0 ? sheet.getRow(startRow - 1) : null;
            writeMatrixAt(sheet, startRow, startCol, rows, templateRow);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }

            // Print where we appended (useful for debugging / future wrapper enhancements)
            System.out.println(startRow);
        }
    }

    /**
     * Find the first empty row by checking the anchor column from top (row 0) downward.
     *
     * Empty means:
     * - row is null, or
     * - anchor cell is null, blank, or empty-string.
     */
    private static int findFirstEmptyRowByAnchorColumn(Sheet sheet, int anchorColIndex) {
        // If there are no physically defined rows, append at row 0.
        if (sheet.getPhysicalNumberOfRows() == 0) {
            return 0;
        }

        int lastRowNum = sheet.getLastRowNum();
        for (int r = 0; r <= lastRowNum; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                return r;
            }
            Cell c = row.getCell(anchorColIndex);
            if (isCellEmpty(c)) {
                return r;
            }
        }

        // No empty row found up to lastRowNum, append after last.
        return lastRowNum + 1;
    }

    private static boolean isCellEmpty(Cell c) {
        if (c == null) {
            return true;
        }
        switch (c.getCellType()) {
            case BLANK:
                return true;
            case STRING:
                String s = c.getStringCellValue();
                return s == null || s.trim().isEmpty();
            default:
                return false;
        }
    }

    /**
     * Writes matrix-like JSON to the sheet starting at (startRow, startCol).
     *
     * When appending, we try to keep the visible look by copying styles from a template row
     * (typically the row immediately above the append position).
     */
    private static void writeMatrixAt(Sheet sheet, int startRow, int startCol, JsonArray rows, Row templateRow) {
        for (int r = 0; r < rows.size(); r++) {
            JsonArray rowArray = rows.get(r).getAsJsonArray();
            int rowIndex = startRow + r;

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
                // Best-effort: copy row height from template row if available.
                if (templateRow != null) {
                    row.setHeight(templateRow.getHeight());
                }
            }

            for (int c = 0; c < rowArray.size(); c++) {
                JsonElement cellElement = rowArray.get(c);
                int colIndex = startCol + c;

                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);

                    // 1) Prefer copying style from template row cell (Excel-like behavior).
                    if (templateRow != null) {
                        Cell tpl = templateRow.getCell(colIndex);
                        if (tpl != null) {
                            CellStyle tplStyle = tpl.getCellStyle();
                            if (tplStyle != null) {
                                cell.setCellStyle(tplStyle);
                            }
                        }
                    }

                    // 2) Fallback: best-effort style inheritance (above/left/column/row).
                    if (cell.getCellStyle() == null || cell.getCellStyle().getIndex() == 0) {
                        applyBestEffortStyle(sheet, rowIndex, colIndex, cell);
                    }
                }

                writeCellValue(cell, cellElement);
            }
        }
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

        // 1) Scan upward (handles cases where the immediate above cell object is null)
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

    private static void writeCellValue(Cell cell, JsonElement cellElement) {
        if (cellElement == null || cellElement.isJsonNull()) {
            // Keep style; only clear content.
            cell.setBlank();
            return;
        }

        if (cellElement.isJsonPrimitive()) {
            if (cellElement.getAsJsonPrimitive().isNumber()) {
                cell.setCellValue(cellElement.getAsDouble());
                return;
            }
            if (cellElement.getAsJsonPrimitive().isBoolean()) {
                cell.setCellValue(cellElement.getAsBoolean());
                return;
            }

            // string
            String s = cellElement.getAsString();
            if (s != null && s.startsWith("=") && s.length() > 1) {
                cell.setCellFormula(s.substring(1));
            } else {
                cell.setCellValue(s);
            }
            return;
        }

        // Fallback: store complex JSON as string
        cell.setCellValue(cellElement.toString());
    }
}