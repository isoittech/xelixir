/**
 * Command line tool that merges a rectangular range of cells.
 * Implements "merge_cells" from README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MergeCellsTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>startCell</li>
     *     <li>endCell</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 4) {
            System.err.println("Usage: MergeCellsTool <filePath> <sheetName> <startCell> <endCell>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String startCell = args[2];
        String endCell = args[3];

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
            sheet.addMergedRegion(range);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }
}
