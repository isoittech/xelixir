/**
 * Command line tool that copies a rectangular range of cells to another
 * location, possibly on another sheet. Implements "copy_range" from
 * README.JA.md.
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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyRangeTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>sourceStart</li>
     *     <li>sourceEnd</li>
     *     <li>targetStart</li>
     *     <li>targetSheet (optional, when omitted the same sheet is used)</li>
     *     <li>copyStyle (optional, true/false; default: true)</li>
     * </ol>
     *
     * Backward compatible parsing:
     * - 5 args: same sheet, copyStyle=true
     * - 6 args: if 6th is boolean => copyStyle=<6th>, targetSheet=sheetName
     *          else targetSheet=<6th>, copyStyle=true
     * - 7 args: targetSheet=<6th>, copyStyle=<7th>
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 5 || args.length > 7) {
            System.err.println("Usage: CopyRangeTool <filePath> <sheetName> <sourceStart> <sourceEnd> <targetStart> [targetSheet] [copyStyle]");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String sourceStart = args[2];
        String sourceEnd = args[3];
        String targetStart = args[4];

        String targetSheetName = sheetName;
        boolean copyStyle = true;

        if (args.length >= 6) {
            String a5 = args[5];
            if ("true".equalsIgnoreCase(a5) || "false".equalsIgnoreCase(a5)) {
                copyStyle = Boolean.parseBoolean(a5);
            } else {
                targetSheetName = a5;
            }
        }
        if (args.length == 7) {
            copyStyle = Boolean.parseBoolean(args[6]);
        }

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sourceSheet = workbook.getSheet(sheetName);
            if (sourceSheet == null) {
                throw new IllegalArgumentException("Source sheet not found: " + sheetName);
            }

            Sheet targetSheet = ExcelUtils.getOrCreateSheet(workbook, targetSheetName);

            CellRangeAddress sourceRange = ExcelRangeUtils.parseRange(sourceStart + ":" + sourceEnd);
            CellAddress targetStartAddr = ExcelRangeUtils.parseCellAddress(targetStart);

            int rowOffset = targetStartAddr.getRow() - sourceRange.getFirstRow();
            int colOffset = targetStartAddr.getColumn() - sourceRange.getFirstColumn();

            for (int r = sourceRange.getFirstRow(); r <= sourceRange.getLastRow(); r++) {
                Row srcRow = sourceSheet.getRow(r);
                Row tgtRow = targetSheet.getRow(r + rowOffset);
                if (tgtRow == null) {
                    tgtRow = targetSheet.createRow(r + rowOffset);
                }

                // Best-effort: copy row height when style copy is enabled.
                if (copyStyle && srcRow != null) {
                    tgtRow.setHeight(srcRow.getHeight());
                }

                for (int c = sourceRange.getFirstColumn(); c <= sourceRange.getLastColumn(); c++) {
                    Cell srcCell = srcRow != null ? srcRow.getCell(c) : null;
                    Cell tgtCell = tgtRow.getCell(c + colOffset);
                    if (tgtCell == null) {
                        tgtCell = tgtRow.createCell(c + colOffset);
                    }
                    if (srcCell != null) {
                        if (copyStyle) {
                            copyCellStyle(srcCell, tgtCell);
                        }
                        copyCellValue(srcCell, tgtCell);
                    } else {
                        tgtCell.setBlank();
                        // If srcCell doesn't exist, we keep target style as-is (do not wipe).
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }

    private static void copyCellStyle(Cell src, Cell dest) {
        if (src == null || dest == null) {
            return;
        }
        try {
            if (src.getCellStyle() != null) {
                dest.setCellStyle(src.getCellStyle());
            }
        } catch (Exception ignored) {
            // ignore
        }
    }

    /**
     * Copies the value and basic type from one cell to another.
     */
    private static void copyCellValue(Cell src, Cell dest) {
        switch (src.getCellType()) {
            case STRING:
                dest.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                dest.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                dest.setCellValue(src.getBooleanCellValue());
                break;
            case FORMULA:
                dest.setCellFormula(src.getCellFormula());
                break;
            default:
                dest.setBlank();
        }
    }
}
