package io.style;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.TextDirection;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import io.TabularDataException;

/**
 * Creates and configures charts on an Excel worksheet using Apache POI.
 *
 * <p>
 * Supports multiple chart types, configurable axes, chart titles, legends and data labels. Custom
 * doughnut chart legends are rendered using worksheet drawing objects.
 * </p>
 *
 * <b>Change log:</b>
 * <ul>
 * <li>6 November 2026 - Initial creation.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.1
 * @since 6 November 2026
 */
public final class WorksheetChartCreator
{
    private final static AxisPosition DEFAULT_XAXIS = AxisPosition.BOTTOM;
    private final static AxisPosition DEFAULT_YAXIS = AxisPosition.LEFT;
    private final XSSFSheet xssfSheet;
    private final XSSFDrawing xssfDrawing;
    private AxisPosition xAxis;
    private AxisPosition yAxis;

    /**
     * Creates a chart creator for the specified worksheet.
     *
     * <p>
     * Initialises the worksheet drawing patriarch and the default axis positions.
     * </p>
     *
     * @param sheet
     *        the worksheet on which charts will be created
     * 
     * @throws NullPointerException
     *         if {@code sheet} is {@code null}
     * @throws TabularDataException
     *         if the worksheet drawing patriarch cannot be created
     */
    public WorksheetChartCreator(XSSFSheet sheet) throws TabularDataException
    {
        this.xssfSheet = Objects.requireNonNull(sheet, "Worksheet must not be null");
        this.xssfDrawing = this.xssfSheet.createDrawingPatriarch();
        this.xAxis = DEFAULT_XAXIS;
        this.yAxis = DEFAULT_YAXIS;
    }

    /**
     * Returns the worksheet associated with this chart creator.
     *
     * @return the associated worksheet
     */
    public XSSFSheet getXssfSheet()
    {
        return xssfSheet;
    }

    /**
     * Overrides the default horizontal axis position for subsequently created charts.
     *
     * @param axis
     *        the new horizontal axis position
     * 
     * @throws NullPointerException
     *         if {@code axis} is {@code null}
     */
    public void overrideDefaultXAxis(AxisPosition axis)
    {
        xAxis = Objects.requireNonNull(axis, "X-axis position must not be null");
    }

    /**
     * Overrides the default vertical axis position for subsequently created charts.
     *
     * @param axis
     *        the new vertical axis position
     * 
     * @throws NullPointerException
     *         if {@code axis} is {@code null}
     */
    public void overrideDefaultYAxis(AxisPosition axis)
    {
        yAxis = Objects.requireNonNull(axis, "Y-axis position must not be null");
    }

    /**
     * Creates and renders the configured charts on the associated worksheet.
     *
     * <p>
     * Each configuration defines the chart type, data series, formatting and worksheet placement.
     * </p>
     *
     * @param configs
     *        the chart configurations to render
     * 
     * @throws NullPointerException
     *         if {@code configs} is {@code null}
     * @throws TabularDataException
     *         if a chart cannot be created or rendered
     */
    public void renderChart(ConfiguredChartData[] configs) throws TabularDataException
    {
        Objects.requireNonNull(configs, "Configurations cannot be null.");

        for (ConfiguredChartData config : configs)
        {
            XDDFChartData data;
            XSSFChart chart = renderChart(config);
            ChartTypes type = config.getChartType();
            XDDFCategoryAxis categoryAxis = null;
            XDDFValueAxis valueAxis = null;

            switch (type)
            {
                case PIE:
                case PIE3D:
                case DOUGHNUT:

                    data = chart.createData(type, null, null);

                    if (data instanceof XDDFDoughnutChartData)
                    {
                        configureDoughnutChart((XDDFDoughnutChartData) data);
                    }

                break;

                default:
                    categoryAxis = chart.createCategoryAxis(xAxis);
                    valueAxis = chart.createValueAxis(yAxis);
                    categoryAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                    valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                    valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
                    data = chart.createData(type, categoryAxis, valueAxis);

                    if (data instanceof XDDFBarChartData)
                    {
                        configureBarChart((XDDFBarChartData) data, valueAxis, config);
                    }

                    else
                    {
                        // TODO add more for other chart types later
                    }
            }

            configureChartTitle(chart, config);
            configureChartLegend(chart, config);

            for (ConfiguredChartData.ChartSeries chartSeries : config.getSeries())
            {
                XDDFCategoryDataSource categories = XDDFDataSourcesFactory.fromArray(chartSeries.getCategories());
                XDDFNumericalDataSource<? extends Number> values = XDDFDataSourcesFactory.fromArray(chartSeries.getValues());
                XDDFChartData.Series series = data.addSeries(categories, values);

                series.setTitle(chartSeries.getLegend(), null);

                if (!(data instanceof XDDFDoughnutChartData))
                {
                    byte[] rgbFillBytes = config.getRgbFillBytes();

                    if (rgbFillBytes != null)
                    {
                        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(rgbFillBytes));
                        XDDFShapeProperties shapeProperties = new XDDFShapeProperties();

                        shapeProperties.setFillProperties(fill);
                        series.setShapeProperties(shapeProperties);
                    }
                }
            }

            chart.plot(data);

            if (data instanceof XDDFBarChartData)
            {
                addDataLabels(chart);
            }

            else if (data instanceof XDDFDoughnutChartData)
            {
                addCustomLegend(chart, config);
            }
        }
    }

    /**
     * Creates a chart on the associated worksheet using the supplied configuration.
     *
     * @param config
     *        the chart configuration
     * @return the newly created chart
     * 
     * @throws TabularDataException
     *         if {@code config} is {@code null}
     */
    public XSSFChart renderChart(ConfiguredChartData config) throws TabularDataException
    {
        if (config == null)
        {
            throw new TabularDataException("Unable to create a new chart. Check the original array for null chart entries");
        }

        XSSFClientAnchor anchor = xssfDrawing.createAnchor(0, 0, 0, 0, config.getFromCol(), config.getFromRow(), config.getToCol(), config.getToRow());
        XSSFChart chart = xssfDrawing.createChart(anchor);

        if (config.getChartTitle() != null)
        {
            chart.setTitleText(config.getChartTitle());
            chart.setTitleOverlay(false);
        }

        return chart;
    }

    // TODO: Add Javadoc
    private static void setBarScaleLogBase(XDDFValueAxis valueAxis, int base)
    {
        if (base > 1)
        {
            valueAxis.setLogBase(base);
            valueAxis.setMinimum(1.0);
        }
    }

    // TODO: Add Javadoc
    private static void configureBarChart(XDDFBarChartData barData, XDDFValueAxis valueAxis, ConfiguredChartData config)
    {
        barData.setBarDirection(config.getBarDirection());
        barData.setBarGrouping(BarGrouping.CLUSTERED);
        barData.setVaryColors(config.getRgbFillBytes() == null);
        barData.setGapWidth(80);

        setBarScaleLogBase(valueAxis, config.getBarScaleLogBase());
    }

    // TODO: Add Javadoc
    private static void configureDoughnutChart(XDDFDoughnutChartData doughnutData)
    {
        doughnutData.setVaryColors(true);
        doughnutData.setHoleSize(60);
    }

    /**
     * Configures the font size and optional font colour of a chart title.
     *
     * <p>
     * Title formatting is applied only when a positive font size is specified. If an RGB colour is
     * supplied, it is applied alongside the font size.
     * </p>
     *
     * @param chart
     *        the chart whose title is to be configured
     * @param config
     *        the chart configuration containing the title formatting properties
     *
     * @throws NullPointerException
     *         if {@code chart} or {@code config} is {@code null}
     * @throws IllegalArgumentException
     *         if the configured RGB colour does not contain exactly three bytes
     */
    private static void configureChartTitle(XSSFChart chart, ConfiguredChartData config)
    {
        byte[] rgb = config.getRgbTitleBytes();
        int fontSize = config.getChartTitleFontPoint();

        if (rgb != null && rgb.length != 3)
        {
            throw new IllegalArgumentException("RGB colour must contain exactly three bytes");
        }

        if (fontSize > 0)
        {
            CTTitle title = chart.getCTChart().getTitle();

            if (title != null && title.getTx() != null && title.getTx().getRich() != null)
            {
                for (CTTextParagraph paragraph : title.getTx().getRich().getPArray())
                {
                    if (fontSize > 0)
                    {
                        if (paragraph.getPPr() == null)
                        {
                            paragraph.addNewPPr();
                        }

                        if (paragraph.getPPr().getDefRPr() == null)
                        {
                            paragraph.getPPr().addNewDefRPr();
                        }

                        paragraph.getPPr().getDefRPr().setSz(fontSize * 100);
                    }

                    for (CTRegularTextRun run : paragraph.getRArray())
                    {
                        CTTextCharacterProperties properties = run.getRPr();

                        if (properties == null)
                        {
                            properties = run.addNewRPr();
                        }

                        if (fontSize > 0)
                        {
                            properties.setSz(fontSize * 100);
                        }

                        if (rgb != null)
                        {
                            if (properties.isSetSolidFill())
                            {
                                properties.unsetSolidFill();
                            }

                            properties.addNewSolidFill().addNewSrgbClr().setVal(rgb);
                        }
                    }
                }
            }
        }
    }

    /**
     * Configures the chart legend position and optional font size.
     *
     * <p>
     * The legend is positioned below the chart. Its text font size is overridden only when a
     * positive font size is specified in the chart configuration.
     * </p>
     *
     * @param chart
     *        the chart whose legend is to be configured
     * @param config
     *        the chart configuration containing the legend formatting properties
     */
    private static void configureChartLegend(XSSFChart chart, ConfiguredChartData config)
    {
        XDDFChartLegend legend = chart.getOrAddLegend();
        CTLegend ctLegend = chart.getCTChart().getLegend();
        CTTextBody textProperties = ctLegend.getTxPr();

        legend.setPosition(LegendPosition.BOTTOM);

        if (textProperties == null)
        {
            textProperties = ctLegend.addNewTxPr();
            textProperties.addNewBodyPr();
            textProperties.addNewLstStyle();
            textProperties.addNewP();
        }

        for (CTTextParagraph paragraph : textProperties.getPArray())
        {
            if (paragraph.getPPr() == null)
            {
                paragraph.addNewPPr();
            }

            if (paragraph.getPPr().getDefRPr() == null)
            {
                paragraph.getPPr().addNewDefRPr();
            }

            if (config.getLegendFontPoint() > 0)
            {
                paragraph.getPPr().getDefRPr().setSz(config.getLegendFontPoint() * 100);
                // paragraph.getPPr().getDefRPr().setB(true);
            }
        }
    }

    // TODO: Add Javadoc
    private static void addDataLabels(XSSFChart chart)
    {
        CTBarChart barChart = chart.getCTChart().getPlotArea().getBarChartArray(0);

        for (CTBarSer barSeries : barChart.getSerArray())
        {
            CTDLbls dataLabels = barSeries.addNewDLbls();

            dataLabels.addNewShowLegendKey().setVal(false);
            dataLabels.addNewShowVal().setVal(true);
            dataLabels.addNewShowCatName().setVal(false);
            dataLabels.addNewShowSerName().setVal(false);
            dataLabels.addNewDLblPos().setVal(STDLblPos.OUT_END);
        }
    }

    /**
     * Adds a custom legend to a doughnut chart.
     *
     * <p>
     * Creates coloured swatches, category labels and calculated percentages using worksheet drawing
     * objects. Also applies the configured colours to the doughnut chart slices and adjusts the
     * plot area to accommodate the custom legend.
     * </p>
     *
     * @param chart
     *        the doughnut chart to configure
     * @param config
     *        the chart configuration containing the legend entries
     */
    private void addCustomLegend(XSSFChart chart, ConfiguredChartData config)
    {
        ConfiguredLegendData legendData = config.getLegendData();

        if (legendData != null && !legendData.getEntries().isEmpty())
        {
            List<double[]> seriesPercentages = new ArrayList<double[]>();
            CTDoughnutChart doughnutChart = chart.getCTChart().getPlotArea().getDoughnutChartArray(0);

            for (CTPieSer series : doughnutChart.getSerArray())
            {
                CTNumDataSource valueSource = series.getVal();

                if (!valueSource.isSetNumLit())
                {
                    continue;
                }

                int index = 0;
                CTNumData values = valueSource.getNumLit();
                double[] percentages = new double[values.sizeOfPtArray()];
                double total = 0.0;

                for (CTNumVal point : values.getPtArray())
                {
                    total += Double.parseDouble(point.getV());
                }

                for (CTNumVal point : values.getPtArray())
                {
                    double value = Double.parseDouble(point.getV());

                    percentages[index++] = (total == 0.0 ? 0.0 : value / total * 100.0);
                }

                seriesPercentages.add(percentages);

                if (series.isSetDLbls())
                {
                    series.unsetDLbls();
                }

                for (int i = 0; i < legendData.getEntries().size(); i++)
                {
                    ConfiguredLegendData.LegendEntry entry = legendData.getEntries().get(i);

                    byte[] rgb = entry.getRgbBytes();

                    if (rgb == null || rgb.length != 3)
                    {
                        continue;
                    }

                    CTDPt dataPoint = series.addNewDPt();

                    dataPoint.addNewIdx().setVal(i);
                    dataPoint.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(rgb);
                }
            }

            if (seriesPercentages.isEmpty())
            {
                return;
            }

            if (chart.getCTChart().isSetLegend())
            {
                chart.getCTChart().unsetLegend();
            }

            int startCol = config.getFromCol() + 5;
            int startRow = config.getFromRow() + 6;

            for (int i = 0; i < legendData.getEntries().size(); i++)
            {
                ConfiguredLegendData.LegendEntry entry = legendData.getEntries().get(i);

                int row = startRow + (i * 2);
                byte[] rgb = entry.getRgbBytes();

                // Colour square
                XSSFClientAnchor colourAnchor = new XSSFClientAnchor(20000, 30000, 90000, 70000, startCol, row, startCol, row + 1);
                XSSFSimpleShape colourSquare = xssfDrawing.createSimpleShape(colourAnchor);

                colourSquare.setShapeType(ShapeTypes.RECT);

                if (rgb != null && rgb.length == 3)
                {
                    colourSquare.setFillColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
                    colourSquare.setLineStyleColor(0, 0, 0);
                }

                // Category label
                XSSFClientAnchor labelAnchor = new XSSFClientAnchor(40000, 0, 0, 0, startCol, row, startCol + 2, row + 1);
                XSSFTextBox labelBox = xssfDrawing.createTextbox(labelAnchor);

                labelBox.setText(entry.getLabel());

                // Percentage columns, one for each doughnut series
                for (int seriesIndex = 0; seriesIndex < seriesPercentages.size(); seriesIndex++)
                {
                    double[] percentages = seriesPercentages.get(seriesIndex);

                    if (i >= percentages.length)
                    {
                        continue;
                    }

                    int percentageCol = startCol + 2 + seriesIndex;

                    XSSFClientAnchor percentageAnchor = new XSSFClientAnchor(0, 0, 0, 0, percentageCol, row, percentageCol + 1, row + 1);
                    XSSFTextBox percentageBox = xssfDrawing.createTextbox(percentageAnchor);
                    percentageBox.setText(String.format(Locale.ROOT, "%.1f%%", percentages[i]));
                    percentageBox.setTextDirection(TextDirection.HORIZONTAL);
                }
            }

            CTPlotArea plotArea = chart.getCTChart().getPlotArea();
            CTLayout layout = plotArea.isSetLayout() ? plotArea.getLayout() : plotArea.addNewLayout();

            /*
             * From raw chart.xml
             *
             * <c:manualLayout>
             * <c:layoutTarget val="inner"/>
             * <c:xMode val="edge"/>
             * <c:yMode val="edge"/>
             * <c:x val="5.4404117454068243E-2"/>
             * <c:y val="0.1295334645669291"/>
             * <c:w val="0.49714587434383201"/>
             * <c:h val="0.79543339895013121"/>
             * </c:manualLayout>
             */

            CTManualLayout manualLayout = layout.addNewManualLayout();

            manualLayout.addNewLayoutTarget().setVal(STLayoutTarget.INNER);
            manualLayout.addNewXMode().setVal(STLayoutMode.EDGE);
            manualLayout.addNewYMode().setVal(STLayoutMode.EDGE);
            manualLayout.addNewX().setVal(0.054404117454068243);
            manualLayout.addNewY().setVal(0.1295334645669291);
            manualLayout.addNewW().setVal(0.49714587434383201);
            manualLayout.addNewH().setVal(0.79543339895013121);
        }
    }

    private void addCustomLegend3(XSSFChart chart, ConfiguredChartData config)
    {
        ConfiguredLegendData legendData = config.getLegendData();

        if (legendData != null && !legendData.getEntries().isEmpty())
        {
            CTDoughnutChart doughnutChart = chart.getCTChart().getPlotArea().getDoughnutChartArray(0);

            for (CTPieSer series : doughnutChart.getSerArray())
            {
                CTNumDataSource valueSource = series.getVal();
                CTNumData values = valueSource.getNumLit();
                double[] percentages = new double[values.sizeOfPtArray()];
                int index = 0;

                double total = 0.0;

                for (CTNumVal point : values.getPtArray())
                {
                    total += Double.parseDouble(point.getV());
                }

                for (CTNumVal point : values.getPtArray())
                {
                    double value = Double.parseDouble(point.getV());
                    percentages[index++] = total == 0.0 ? 0.0 : value / total * 100.0;
                    System.out.printf("LOOK: %s\n", percentages[index - 1]);
                }

                if (series.isSetDLbls())
                {
                    series.unsetDLbls();
                }

                for (int i = 0; i < legendData.getEntries().size(); i++)
                {
                    ConfiguredLegendData.LegendEntry entry = legendData.getEntries().get(i);

                    byte[] rgb = entry.getRgbBytes();

                    if (rgb == null || rgb.length != 3)
                    {
                        continue;
                    }

                    CTDPt dataPoint = series.addNewDPt();

                    dataPoint.addNewIdx().setVal(i);
                    dataPoint.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(rgb);
                }
            }

            if (chart.getCTChart().isSetLegend())
            {
                chart.getCTChart().unsetLegend();
            }

            int rowOffset = 0;
            int startCol = config.getFromCol() + 5;
            int startRow = config.getFromRow() + 6;

            for (ConfiguredLegendData.LegendEntry entry : legendData.getEntries())
            {
                byte[] rgb = entry.getRgbBytes();
                int row = startRow + (rowOffset * 2);

                XSSFClientAnchor colourAnchor = new XSSFClientAnchor(20000, 30000, 90000, 70000, startCol, row, startCol, row + 1);
                XSSFSimpleShape colourSquare = xssfDrawing.createSimpleShape(colourAnchor);

                colourSquare.setShapeType(ShapeTypes.RECT);

                if (rgb != null && rgb.length == 3)
                {
                    colourSquare.setFillColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
                    colourSquare.setLineStyleColor(0, 0, 0);
                }

                // Category label
                XSSFClientAnchor labelAnchor = new XSSFClientAnchor(40000, 0, 0, 0, startCol, row, startCol + 2, row + 1);
                XSSFTextBox labelBox = xssfDrawing.createTextbox(labelAnchor);

                labelBox.setText(entry.getLabel());

                // Percentage
                XSSFClientAnchor percentageAnchor = new XSSFClientAnchor(0, 0, 0, 0, startCol + 2, row, startCol + 3, row + 1);
                XSSFTextBox percentageBox = xssfDrawing.createTextbox(percentageAnchor);

                percentageBox.setText(entry.getPercentageText());
                percentageBox.setTextDirection(TextDirection.HORIZONTAL);

                rowOffset++;
            }

            CTPlotArea plotArea = chart.getCTChart().getPlotArea();
            CTLayout layout = (plotArea.isSetLayout() ? plotArea.getLayout() : plotArea.addNewLayout());

            /*
             * From raw chart.xml
             * <c:manualLayout>
             * <c:layoutTarget val="inner"/>
             * <c:xMode val="edge"/>
             * <c:yMode val="edge"/>
             * <c:x val="5.4404117454068243E-2"/>
             * <c:y val="0.1295334645669291"/>
             * <c:w val="0.49714587434383201"/>
             * <c:h val="0.79543339895013121"/>
             * </c:manualLayout>
             */
            CTManualLayout manualLayout = layout.addNewManualLayout();

            manualLayout.addNewLayoutTarget().setVal(STLayoutTarget.INNER);
            manualLayout.addNewXMode().setVal(STLayoutMode.EDGE);
            manualLayout.addNewYMode().setVal(STLayoutMode.EDGE);
            manualLayout.addNewX().setVal(0.054404117454068243);
            manualLayout.addNewY().setVal(0.1295334645669291);
            manualLayout.addNewW().setVal(0.49714587434383201);
            manualLayout.addNewH().setVal(0.79543339895013121);
        }
    }
}