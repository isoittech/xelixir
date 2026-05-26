/**
 * Command line tool that validates the syntax of an Excel formula by
 * attempting to assign it to a temporary cell. If the formula is
 * syntactically invalid, POI will throw an exception.
 *
 * <p>Implements "validate_formula_syntax" from README.JA.md.</p>
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ValidateFormulaSyntaxTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath - path to the .xlsx file</li>
     *     <li>sheetName - worksheet name</li>
     *     <li>formula - Excel formula</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("Usage: ValidateFormulaSyntaxTool <filePath> <sheetName> <formula>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String formula = args[2];

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

            // Use a temporary cell to test the formula syntax.
            Row row = sheet.getRow(0);
            if (row == null) {
                row = sheet.createRow(0);
            }
            Cell cell = row.getCell(0);
            if (cell == null) {
                cell = row.createCell(0);
            }

            if (!formula.startsWith("=")) {
                formula = "=" + formula;
            }

            // If formula is invalid, this call may throw an exception.
            cell.setCellFormula(formula.substring(1));

            // If we reach this point, the syntax is considered valid.
            System.out.println("OK");
        }
    }
}
