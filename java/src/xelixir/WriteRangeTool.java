/**
 * Command line tool that writes a rectangular block of values to a worksheet,
 * starting at a given top-left cell (A1 notation).
 *
 * Arguments:
 * <ol>
 *     <li>filePath - path to the .xlsx file</li>
 *     <li>sheetName - name of the worksheet</li>
 *     <li>startCell - top-left cell address (e.g. "B3")</li>
 *     <li>jsonData - JSON array of arrays representing rows and cells</li>
 * </ol>
 *
 * Notes:
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
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteRangeTool {

    public static void main(String[] args) throws Exception {
        if (args.length != 4) {
            System.err.println("Usage: WriteRangeTool <filePath> <sheetName> <startCell> <jsonData>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String startCellStr = args[2];
        String jsonData = args[3];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = ExcelUtils.getOrCreateSheet(workbook, sheetName);

            CellAddress startAddr = ExcelRangeUtils.parseCellAddress(startCellStr);
            int startRow = startAddr.getRow();
            int startCol = startAddr.getColumn();

            JsonArray rows = JsonParser.parseString(jsonData).getAsJsonArray();
            writeMatrixAt(sheet, startRow, startCol, rows);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * Writes matrix-like JSON to the sheet starting at (startRow, startCol).
     */
    private static void writeMatrixAt(Sheet sheet, int startRow, int startCol, JsonArray rows) {
        for (int r = 0; r < rows.size(); r++) {
            JsonArray rowArray = rows.get(r).getAsJsonArray();
            int rowIndex = startRow + r;

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            for (int c = 0; c < rowArray.size(); c++) {
                JsonElement cellElement = rowArray.get(c);
                int colIndex = startCol + c;

                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                    // Best-effort style inheritance for newly created cells.
                    applyBestEffortStyle(sheet, rowIndex, colIndex, cell);
                }

                writeCellValue(cell, cellElement);
            }
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
}