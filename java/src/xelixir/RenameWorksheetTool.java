/**
 * Command line tool that renames a worksheet in an existing workbook.
 * Implements the "rename_worksheet" tool from README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RenameWorksheetTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath - path to the .xlsx file</li>
     *     <li>oldName - existing sheet name</li>
     *     <li>newName - new sheet name</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("Usage: RenameWorksheetTool <filePath> <oldName> <newName>");
            System.exit(1);
        }

        String filePath = args[0];
        String oldName = args[1];
        String newName = args[2];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            int index = workbook.getSheetIndex(oldName);
            if (index < 0) {
                throw new IllegalArgumentException("Sheet not found: " + oldName);
            }

            workbook.setSheetName(index, newName);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }
}
