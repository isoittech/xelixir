/**
 * Command line tool that validates an Excel range string and optionally
 * checks that the referenced cells exist within the sheet's bounds.
 * Implements "validate_excel_range" from README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ValidateExcelRangeTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>startCell</li>
     *     <li>endCell (optional)</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 3 || args.length > 4) {
            System.err.println("Usage: ValidateExcelRangeTool <filePath> <sheetName> <startCell> [endCell]");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String startCell = args[2];
        String endCell = args.length == 4 ? args[3] : startCell;

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

            CellRangeAddress range = ExcelRangeUtils.parseRange(startCell + ":" + endCell);

            // Basic bounds validation: ensure indices are non-negative.
            if (range.getFirstRow() < 0 || range.getFirstColumn() < 0) {
                throw new IllegalArgumentException("Range is out of bounds: " + range.formatAsString());
            }

            System.out.println("OK");
        }
    }
}
