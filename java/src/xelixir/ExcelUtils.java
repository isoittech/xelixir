/**
 * Common utility functions for working with Excel workbooks and worksheets.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class ExcelUtils {

    private ExcelUtils() {
    }

    /**
     * Opens an existing Excel workbook (.xlsx).
     *
     * @param filePath absolute path to the workbook file
     * @return open {@link Workbook} instance
     * @throws IOException if the file does not exist or cannot be read
     */
    public static Workbook openWorkbook(String filePath) throws IOException {
        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("Workbook not found: " + filePath);
        }
        FileInputStream fis = new FileInputStream(file);
        return new XSSFWorkbook(fis);
    }

    /**
     * Creates a new workbook with a single worksheet.
     *
     * @param filePath  absolute path of the file to create
     * @param sheetName name of the first sheet; if {@code null} or empty,
     *                  "Sheet1" is used
     * @throws IOException if the workbook cannot be written
     */
    public static void createWorkbook(String filePath, String sheetName) throws IOException {
        String name = (sheetName == null || sheetName.isEmpty()) ? "Sheet1" : sheetName;
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(new File(filePath))) {
            workbook.createSheet(name);
            workbook.write(fos);
        }
    }

    /**
     * Gets an existing sheet by name or creates it if it does not exist.
     *
     * @param workbook workbook to search in
     * @param sheetName sheet name
     * @return existing or newly created {@link Sheet}
     */
    public static Sheet getOrCreateSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }
        return sheet;
    }
}
