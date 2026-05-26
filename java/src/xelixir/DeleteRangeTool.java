/**
 * Command line tool that deletes a rectangular cell range and shifts
 * remaining cells either up or left. Implements "delete_range" from
 * README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DeleteRangeTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>startCell</li>
     *     <li>endCell</li>
     *     <li>shiftDirection ("up" or "left")</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 5) {
            System.err.println("Usage: DeleteRangeTool <filePath> <sheetName> <startCell> <endCell> <shiftDirection>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String startCell = args[2];
        String endCell = args[3];
        String shiftDirection = args[4];

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

            if ("up".equalsIgnoreCase(shiftDirection)) {
                sheet.shiftRows(range.getLastRow() + 1, sheet.getLastRowNum(),
                        range.getFirstRow() - range.getLastRow() - 1);
            } else if ("left".equalsIgnoreCase(shiftDirection)) {
                // For column shift, use a CellCopyPolicy to copy remaining cells
                // over the deleted area. This is a simplified implementation.
                for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
                    sheet.shiftColumns(range.getLastColumn() + 1, sheet.getRow(r).getLastCellNum(),
                            range.getFirstColumn() - range.getLastColumn() - 1);
                }
            } else {
                throw new IllegalArgumentException("shiftDirection must be 'up' or 'left'");
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }
}
