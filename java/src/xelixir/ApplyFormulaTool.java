/**
 * Command line tool that sets a formula into a specific cell.
 * Implements the "apply_formula" tool from README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApplyFormulaTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath - path to the .xlsx file</li>
     *     <li>sheetName - worksheet name</li>
     *     <li>cell - target cell address (for example "C1")</li>
     *     <li>formula - Excel formula string (for example "=SUM(A1:B1)")</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 4) {
            System.err.println("Usage: ApplyFormulaTool <filePath> <sheetName> <cell> <formula>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String cellAddressStr = args[2];
        String formula = args[3];

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

            CellAddress addr = ExcelRangeUtils.parseCellAddress(cellAddressStr);
            Row row = sheet.getRow(addr.getRow());
            if (row == null) {
                row = sheet.createRow(addr.getRow());
            }
            Cell cell = row.getCell(addr.getColumn());
            if (cell == null) {
                cell = row.createCell(addr.getColumn());
            }

            if (!formula.startsWith("=")) {
                formula = "=" + formula;
            }
            cell.setCellFormula(formula.substring(1));

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }
}
