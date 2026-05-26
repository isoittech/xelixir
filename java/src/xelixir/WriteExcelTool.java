/**
 * Command line tool that writes a rectangular block of values to a
 * worksheet. Data is provided as a simple JSON array of arrays and
 * written starting from the top-left cell.
 *
 * <p>The interface corresponds to the "write_excel" tool described in
 * README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import com.google.gson.Gson;
import com.google.gson.JsonElement;
import com.google.gson.JsonArray;
import com.google.gson.JsonParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelTool {

    /**
     * Entry point. Arguments:
     * <ol>
     *     <li>filePath - path to the .xlsx file</li>
     *     <li>sheetName - name of the worksheet</li>
     *     <li>jsonData - JSON array of arrays representing rows and cells</li>
     * </ol>
     *
     * <p>Example of jsonData:
     * <pre>
     *   [["A1", "B1"], ["A2", "B2"]]
     * </pre>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("Usage: WriteExcelTool <filePath> <sheetName> <jsonData>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String jsonData = args[2];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = ExcelUtils.getOrCreateSheet(workbook, sheetName);

            JsonArray rows = JsonParser.parseString(jsonData).getAsJsonArray();
            writeMatrix(sheet, rows);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * Writes matrix-like JSON to the sheet starting from cell A1.
     */
    private static void writeMatrix(Sheet sheet, JsonArray rows) {
        for (int r = 0; r < rows.size(); r++) {
            JsonArray rowArray = rows.get(r).getAsJsonArray();
            Row row = sheet.getRow(r);
            if (row == null) {
                row = sheet.createRow(r);
            }
            for (int c = 0; c < rowArray.size(); c++) {
                JsonElement cellElement = rowArray.get(c);
                Cell cell = row.getCell(c);
                if (cell == null) {
                    cell = row.createCell(c);
                    // Best-effort style inheritance for newly created cells.
                    applyBestEffortStyle(sheet, r, c, cell);
                }

                if (cellElement.isJsonNull()) {
                    // Keep style; only clear content.
                    cell.setBlank();
                } else if (cellElement.isJsonPrimitive()) {
                    if (cellElement.getAsJsonPrimitive().isNumber()) {
                        cell.setCellValue(cellElement.getAsDouble());
                    } else if (cellElement.getAsJsonPrimitive().isBoolean()) {
                        cell.setCellValue(cellElement.getAsBoolean());
                    } else {
                        cell.setCellValue(cellElement.getAsString());
                    }
                } else {
                    // Fallback: store complex JSON as string
                    cell.setCellValue(cellElement.toString());
                }
            }
        }
    }

    /**
     * Best-effort style inheritance to keep the "Excel visible look" when new cells are created.
     *
     * Priority:
     *  1) Above cell style (same column, previous row)
     *  2) Left cell style (same row, previous column)
     *  3) Column style
     *  4) Row style
     */
    private static void applyBestEffortStyle(Sheet sheet, int rowIndex, int colIndex, Cell target) {
        CellStyle style = null;

        // 1) Above
        if (rowIndex > 0) {
            Row aboveRow = sheet.getRow(rowIndex - 1);
            if (aboveRow != null) {
                Cell above = aboveRow.getCell(colIndex);
                if (above != null) {
                    style = above.getCellStyle();
                }
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
