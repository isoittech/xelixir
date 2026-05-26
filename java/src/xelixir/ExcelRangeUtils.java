/**
 * Utility functions for working with Excel cell addresses and ranges
 * such as "A1" or "A1:C10".
 */
package xelixir;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

public final class ExcelRangeUtils {

    private ExcelRangeUtils() {
    }

    /**
     * Parses a single cell address (for example "C5") into a {@link CellAddress}.
     *
     * @param address cell address in A1 notation
     * @return parsed {@link CellAddress}
     * @throws IllegalArgumentException if the address is invalid
     */
    public static CellAddress parseCellAddress(String address) {
        try {
            return new CellAddress(address);
        } catch (Exception ex) {
            throw new IllegalArgumentException("Invalid cell address: " + address, ex);
        }
    }

    /**
     * Parses a range like "A1:C10" into a {@link CellRangeAddress}.
     *
     * @param range range in A1 notation
     * @return parsed {@link CellRangeAddress}
     * @throws IllegalArgumentException if the range is invalid
     */
    public static CellRangeAddress parseRange(String range) {
        if (range == null || !range.contains(":")) {
            throw new IllegalArgumentException("Invalid range format: " + range);
        }
        String[] parts = range.split(":", 2);
        CellAddress start = parseCellAddress(parts[0]);
        CellAddress end = parseCellAddress(parts[1]);
        return new CellRangeAddress(start.getRow(), end.getRow(), start.getColumn(), end.getColumn());
    }
}
