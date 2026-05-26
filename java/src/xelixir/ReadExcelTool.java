/**
 * Command line tool that reads a rectangular range of cells from
 * an Excel workbook and prints the contents as a JSON array.
 *
 * <p>The command line interface is intentionally simple and matches the
 * structure shown in README.JA.md under the "read_excel" tool:
 *
 * <pre>
 *   java -cp ... xelixir.ReadExcelTool filePath sheetName range
 * </pre>
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

public class ReadExcelTool {

    /**
     * Entry point. Arguments:
     * <ol>
     *     <li>filePath - path to the .xlsx file</li>
     *     <li>sheetName - name of the worksheet</li>
     *     <li>range - A1-style range (for example "A1:C10")</li>
     * </ol>
     * The result is printed as JSON to standard output.
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("Usage: ReadExcelTool <filePath> <sheetName> <range>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String range = args[2];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            CellRangeAddress cellRange = ExcelRangeUtils.parseRange(range);

            Object[][] data = readRangeAsMatrix(sheet, cellRange);

            Gson gson = new GsonBuilder().serializeNulls().create();
            System.out.println(gson.toJson(data));
        }
    }

    /**
     * Reads the given cell range from the sheet and returns a 2D array
     * of values. Cell values are converted to simple Java types
     * (String, Double, Boolean) or {@code null}.
     */
    private static Object[][] readRangeAsMatrix(Sheet sheet, CellRangeAddress range) {
        int rows = range.getLastRow() - range.getFirstRow() + 1;
        int cols = range.getLastColumn() - range.getFirstColumn() + 1;
        Object[][] matrix = new Object[rows][cols];

        for (int r = 0; r < rows; r++) {
            Row row = sheet.getRow(range.getFirstRow() + r);
            for (int c = 0; c < cols; c++) {
                Object value = null;
                if (row != null) {
                    Cell cell = row.getCell(range.getFirstColumn() + c);
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                value = cell.getStringCellValue();
                                break;
                            case NUMERIC:
                                value = cell.getNumericCellValue();
                                break;
                            case BOOLEAN:
                                value = cell.getBooleanCellValue();
                                break;
                            case FORMULA:
                                // keep formula as string so MCP client can decide how to handle it
                                value = "=" + cell.getCellFormula();
                                break;
                            default:
                                value = null;
                        }
                    }
                }
                matrix[r][c] = value;
            }
        }
        return matrix;
    }
}
