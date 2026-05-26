/**
 * Command line tool that creates a new worksheet in an existing workbook.
 * This corresponds to the "create_sheet" tool in README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateSheetTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath - path to the .xlsx file</li>
     *     <li>sheetName - name of the sheet to create</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 2) {
            System.err.println("Usage: CreateSheetTool <filePath> <sheetName>");
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

            Sheet existing = workbook.getSheet(sheetName);
            if (existing != null) {
                throw new IllegalArgumentException("Sheet already exists: " + sheetName);
            }

            workbook.createSheet(sheetName);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }
}
