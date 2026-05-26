/**
 * Command line tool that creates a simple chart on a worksheet using
 * a given data range. Implements "create_chart" from README.JA.md.
 */
package xelixir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xddf.usermodel.chart.*;

public class CreateChartTool {

    /**
     * Arguments (string values):
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>dataRange (e.g. "A1:C10")</li>
     *     <li>chartType (column, line, bar, area, scatter, pie)</li>
     *     <li>targetCell (e.g. "E1")</li>
     *     <li>title (optional)</li>
     *     <li>xAxis (optional)</li>
     *     <li>yAxis (optional)</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 5 || args.length > 8) {
            System.err.println("Usage: CreateChartTool <filePath> <sheetName> <dataRange> <chartType> <targetCell> [title] [xAxis] [yAxis]");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String dataRange = args[2];
        String chartType = args[3];
        String targetCell = args[4];
        String title = args.length >= 6 ? args[5] : null;
        String xAxisTitle = args.length >= 7 ? args[6] : null;
        String yAxisTitle = args.length >= 8 ? args[7] : null;

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            CellRangeAddress range = ExcelRangeUtils.parseRange(dataRange);

            XSSFDrawing drawing = ((org.apache.poi.xssf.usermodel.XSSFSheet) sheet).createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    ExcelRangeUtils.parseCellAddress(targetCell).getColumn(),
                    ExcelRangeUtils.parseCellAddress(targetCell).getRow(),
                    ExcelRangeUtils.parseCellAddress(targetCell).getColumn() + 10,
                    ExcelRangeUtils.parseCellAddress(targetCell).getRow() + 20);
            XSSFChart chart = drawing.createChart(anchor);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);

            if (xAxisTitle != null) {
                bottomAxis.setTitle(xAxisTitle);
            }
            if (yAxisTitle != null) {
                leftAxis.setTitle(yAxisTitle);
            }

            XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange((org.apache.poi.xssf.usermodel.XSSFSheet) sheet, range);
            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange((org.apache.poi.xssf.usermodel.XSSFSheet) sheet, range);

            XDDFChartData data;
            ChartTypes type = mapChartType(chartType);
            data = chart.createData(type, bottomAxis, leftAxis);
            XDDFChartData.Series series = data.addSeries(xs, ys);
            series.setTitle(title, null);

            chart.plot(data);

            if (title != null) {
                chart.setTitleText(title);
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }

    private static ChartTypes mapChartType(String chartType) {
        if (chartType == null) {
            return ChartTypes.BAR;
        }
        switch (chartType.toLowerCase()) {
            case "column":
                return ChartTypes.BAR;
            case "line":
                return ChartTypes.LINE;
            case "bar":
                return ChartTypes.BAR;
            case "area":
                return ChartTypes.AREA;
            case "scatter":
                return ChartTypes.SCATTER;
            case "pie":
                return ChartTypes.PIE;
            default:
                return ChartTypes.BAR;
        }
    }
}
