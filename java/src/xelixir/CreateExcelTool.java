/**
 * Command line tool that creates a new workbook (.xlsx) with a single sheet.
 * This is the implementation of the "create_excel" tool described in
 * README.JA.md.
 */
package xelixir;

public class CreateExcelTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath - path of the .xlsx file to create</li>
     *     <li>sheetName - optional name of the first sheet (defaults to "Sheet1")</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 1 || args.length > 2) {
            System.err.println("Usage: CreateExcelTool <filePath> [sheetName]");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args.length == 2 ? args[1] : "Sheet1";

        ExcelUtils.createWorkbook(filePath, sheetName);
    }
}
