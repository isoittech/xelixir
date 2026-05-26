/**
 * Command line tool that lists worksheet names in an Excel workbook and prints them as JSON.
 *
 * <pre>
 *   java -cp ... xelixir.ListSheetsTool filePath
 * </pre>
 */
package xelixir;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.Workbook;

public class ListSheetsTool {

    /**
     * Entry point. Arguments:
     * <ol>
     *   <li>filePath - path to the .xlsx file</li>
     * </ol>
     * The result is printed as JSON array of sheet names to standard output.
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 1) {
            System.err.println("Usage: ListSheetsTool <filePath>");
            System.exit(1);
        }

        String filePath = args[0];

        try (Workbook workbook = ExcelUtils.openWorkbook(filePath)) {
            List<String> names = new ArrayList<>();
            int count = workbook.getNumberOfSheets();
            for (int i = 0; i < count; i++) {
                names.add(workbook.getSheetName(i));
            }
            System.out.println(new Gson().toJson(names));
        } catch (IOException e) {
            throw e;
        }
    }
}
