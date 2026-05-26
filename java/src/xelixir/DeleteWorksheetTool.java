/**
 * Command line tool that deletes a worksheet from an existing workbook.
 * Implements the "delete_worksheet" tool from README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DeleteWorksheetTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath - path to the .xlsx file</li>
     *     <li>sheetName - name of the sheet to delete</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 2) {
            System.err.println("Usage: DeleteWorksheetTool <filePath> <sheetName>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            int index = workbook.getSheetIndex(sheetName);
            if (index < 0) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            workbook.removeSheetAt(index);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }
}
