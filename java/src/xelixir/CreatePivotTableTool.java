/**
 * Command line tool that simulates the creation of a pivot table.
 *
 * <p>Due to limitations of the underlying Excel library, this tool does
 * not create a real pivot table but accepts the arguments and prints a
 * confirmation message. This mirrors the behavior described in
 * README.JA.md.</p>
 */
package xelixir;

public class CreatePivotTableTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>dataRange</li>
     *     <li>rows (comma-separated labels)</li>
     *     <li>values (comma-separated labels)</li>
     *     <li>columns (optional, comma-separated)</li>
     *     <li>aggFunc (for example "sum")</li>
     * </ol>
     */
    public static void main(String[] args) {
        // This tool intentionally does not modify Excel files.
        System.out.println("Pivot table request accepted (no real pivot table created).");
    }
}
