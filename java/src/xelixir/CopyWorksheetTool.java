/**
 * Command line tool that copies a worksheet within a workbook.
 * Implements the "copy_worksheet" tool from README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyWorksheetTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath - path to the .xlsx file</li>
     *     <li>sourceSheet - name of the sheet to copy</li>
     *     <li>targetSheet - name of the new sheet</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("Usage: CopyWorksheetTool <filePath> <sourceSheet> <targetSheet>");
            System.exit(1);
        }

        String filePath = args[0];
        String sourceSheet = args[1];
        String targetSheet = args[2];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            int sourceIndex = workbook.getSheetIndex(sourceSheet);
            if (sourceIndex < 0) {
                throw new IllegalArgumentException("Source sheet not found: " + sourceSheet);
            }

            Sheet cloned = workbook.cloneSheet(sourceIndex);
            int newIndex = workbook.getSheetIndex(cloned.getSheetName());
            workbook.setSheetName(newIndex, targetSheet);

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }
}
