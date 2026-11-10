package io.style;

import java.util.List;
import java.util.Locale;
import java.util.Objects;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFNoFillProperties;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrossBetween;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.ScatterStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFView3D;
import org.apache.poi.xssf.usermodel.TextDirection;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDPt;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLayout;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTManualLayout;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLayoutMode;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLayoutTarget;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import io.TabularDataException;
import io.style.ConfiguredChartData.ChartDataPoint;

/**
 * Creates and configures charts on an Excel worksheet using Apache POI.
 *
 * <p>
 * Supports multiple chart types, configurable axes, chart titles, legends and data labels. Custom
 * pie based chart legends are rendered using worksheet drawing objects.
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
     * Creates and renders multiple configured charts on the associated worksheet.
     *
     * @param configs
     *        the array of chart configurations to render
     * @throws NullPointerException
     *         if {@code configs} is {@code null}
     * @throws TabularDataException
     *         if any chart cannot be created or rendered
     */
    public void renderChart(ConfiguredChartData[] configs) throws TabularDataException
    {
        Objects.requireNonNull(configs, "Configurations cannot be null.");

        for (ConfiguredChartData config : configs)
        {
            renderChart(config);
        }
    }

    /**
     * Renders a single chart on the associated worksheet using the specified configuration.
     *
     * @param config
     *        the chart configuration
     * @return the newly rendered chart
     *
     * @throws TabularDataException
     *         if {@code config} is {@code null}
     */
    public XSSFChart renderChart(ConfiguredChartData config) throws TabularDataException
    {
        if (config == null)
        {
            throw new TabularDataException("Unable to create a new chart. Check the configuration entry for null values.");
        }

        XDDFChartData data;
        boolean isPieBased = false;
        List<ChartDataPoint> dataPoint = config.getSeries();
        int len = dataPoint.size();
        String[] catArr = new String[len];
        Number[] valArr = new Number[len];
        ChartTypes type = config.getChartType();
        XDDFCategoryAxis categoryAxis = null;
        XDDFValueAxis valueAxis = null;

        XSSFClientAnchor anchor = xssfDrawing.createAnchor(0, 0, 0, 0, config.getFromCol(), config.getFromRow(), config.getToCol(), config.getToRow());
        XSSFChart chart = xssfDrawing.createChart(anchor);

        for (int i = 0; i < len; i++)
        {
            catArr[i] = dataPoint.get(i).getCategory();
            valArr[i] = dataPoint.get(i).getValue();
        }

        if (type == ChartTypes.BAR3D || type == ChartTypes.LINE3D || type == ChartTypes.PIE3D)
        {
            configure3DView(chart);
        }

        switch (type)
        {
            case PIE:
            case PIE3D:
            case DOUGHNUT:
                isPieBased = true;
                data = chart.createData(type, null, null);

                if (data instanceof XDDFDoughnutChartData)
                {
                    ((XDDFDoughnutChartData) data).setHoleSize(60);
                }

                else if (data instanceof XDDFPie3DChartData)
                {
                    configure3DView(chart);
                }
            break;

            default:
                categoryAxis = chart.createCategoryAxis(xAxis);
                valueAxis = chart.createValueAxis(yAxis);
                categoryAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
                data = chart.createData(type, categoryAxis, valueAxis);
        }

        XDDFCategoryDataSource categories = XDDFDataSourcesFactory.fromArray(catArr);
        XDDFNumericalDataSource<? extends Number> values = XDDFDataSourcesFactory.fromArray(valArr);
        XDDFChartData.Series series = data.addSeries(categories, values);

        configureChartTitle(chart, config);

        if (data instanceof XDDFBarChartData)
        {
            configureBarChart((XDDFBarChartData) data, valueAxis, config);
        }

        else if (data instanceof XDDFLineChartData)
        {
            configureLineChart((XDDFLineChartData) data, config);
        }

        else if (data instanceof XDDFAreaChartData)
        {
            ((XDDFAreaChartData) data).setVaryColors(false);
        }

        if (!isPieBased)
        {
            configureChartLegend(chart, config);
            series.setTitle(config.getLegendLabel(), null);

            byte[] rgbFillBytes = config.getRgbFillBytes();

            if (rgbFillBytes != null || data instanceof XDDFScatterChartData)
            {
                XDDFShapeProperties shapeProperties = new XDDFShapeProperties();

                if (rgbFillBytes != null)
                {
                    XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(rgbFillBytes));

                    if (data instanceof XDDFLineChartData)
                    {
                        XDDFLineProperties lineProperties = new XDDFLineProperties();
                        lineProperties.setFillProperties(fill);
                        lineProperties.setWidth(2.5);
                        shapeProperties.setLineProperties(lineProperties);
                    }

                    else if (data instanceof XDDFAreaChartData)
                    {
                        XDDFLineProperties lineProperties = new XDDFLineProperties();
                        lineProperties.setFillProperties(fill);
                        lineProperties.setWidth(1.5);
                        shapeProperties.setLineProperties(lineProperties);
                        shapeProperties.setFillProperties(fill);
                    }

                    else
                    {
                        shapeProperties.setFillProperties(fill);
                    }
                }

                if (data instanceof XDDFScatterChartData)
                {
                    XDDFLineProperties lineProperties = new XDDFLineProperties();
                    lineProperties.setFillProperties(new XDDFNoFillProperties());
                    shapeProperties.setLineProperties(lineProperties);
                }

                series.setShapeProperties(shapeProperties);
            }
        }

        if (data instanceof XDDFScatterChartData)
        {
            configureScatterChart((XDDFScatterChartData) data, config);
        }

        chart.plot(data);

        if (isPieBased)
        {
            drawCustomLegendBox(chart, config);
        }

        else if (data instanceof XDDFBarChartData)
        {
            addDataLabels(chart);
        }

        return chart;
    }

    private void configure3DView(XSSFChart chart)
    {
        XDDFView3D view3D = chart.getOrAddView3D();

        view3D.setXRotationAngle((byte) 20);
        view3D.setYRotationAngle(20);
        view3D.setPerspectiveAngle((short) 30);
        view3D.setRightAngleAxes(false);
        view3D.setDepthPercent(150);
    }

    /**
     * Configures the font size and optional font colour of a chart title.
     *
     * <p>
     * Font size and font colour are applied independently. Only a positive font size is applied.
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
    private void configureChartTitle(XSSFChart chart, ConfiguredChartData config)
    {
        byte[] rgb = config.getRgbTitleBytes();
        int fontSize = config.getChartTitleFontPoint();

        if (config.getChartTitle() != null)
        {
            chart.setTitleText(config.getChartTitle());
            chart.setTitleOverlay(false);
        }

        if (fontSize > 0 || rgb != null)
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
    private void configureChartLegend(XSSFChart chart, ConfiguredChartData config)
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

    /**
     * Adds a custom legend to a pie based chart.
     *
     * <p>
     * Creates coloured swatches, category labels and calculated percentages using worksheet drawing
     * objects. Also applies the configured colours to the chart slices and adjusts the plot area to
     * accommodate the custom legend.
     * </p>
     *
     * <p>
     * This method currently supports one series per chart. Support for multiple series may be
     * introduced in a future version.
     * </p>
     *
     * @param chart
     *        the pie based chart to configure
     * @param config
     *        the chart configuration containing the category, value and colour data points
     */
    private void drawCustomLegendBox(XSSFChart chart, ConfiguredChartData config)
    {
        double total = 0.0;
        CTPieSer[] seriesArray;
        List<ChartDataPoint> dataPoints = config.getSeries();

        switch (config.getChartType())
        {
            case DOUGHNUT:
                seriesArray = chart.getCTChart().getPlotArea().getDoughnutChartArray(0).getSerArray();
            break;

            case PIE:
                seriesArray = chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray();
            break;

            case PIE3D:
                seriesArray = chart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray();
            break;

            default:
                throw new IllegalArgumentException("Unsupported chart type.");
        }

        // We are interested in exactly one series at this stage.
        CTPieSer series = seriesArray[0];

        // As the configuration is the source of truth, it calculates the total percentage first.
        for (int i = 0; i < dataPoints.size(); i++)
        {
            ChartDataPoint point = Objects.requireNonNull(dataPoints.get(i), "Chart data point must not be null at index " + i);
            double value = point.getValue().doubleValue();

            if (Double.isNaN(value) || Double.isInfinite(value) || value < 0.0)
            {
                throw new IllegalStateException("Invalid pie chart value at index " + i + ": " + value);
            }

            total += value;
        }

        if (Double.isInfinite(total))
        {
            throw new IllegalStateException("Total of pie chart values exceeds the supported range.");
        }

        // Remove the native chart legend. The custom worksheet-based legend replaces it.
        if (chart.getCTChart().isSetLegend())
        {
            chart.getCTChart().unsetLegend();
        }

        // Remove native data labels. Percentages are displayed in the custom legend box instead.
        if (series.isSetDLbls())
        {
            series.unsetDLbls();
        }

        while (series.sizeOfDPtArray() > 0)
        {
            series.removeDPt(0);
        }

        for (int i = 0; i < dataPoints.size(); i++)
        {
            byte[] rgb = dataPoints.get(i).getColourRGB();

            if (rgb != null)
            {
                CTDPt dataPoint = series.addNewDPt();
                dataPoint.addNewIdx().setVal(i);
                dataPoint.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(rgb);
            }
        }

        /*
         * Render the custom legend using worksheet drawing objects.
         */
        int offsetCol = config.getFromCol() + 5;
        int offsetRow = config.getFromRow() + 6;

        for (int i = 0; i < dataPoints.size(); i++)
        {
            ChartDataPoint point = dataPoints.get(i);
            int row = offsetRow + (i * 2);
            byte[] rgb = point.getColourRGB();

            // Colour square, positioned to the left of the category label.
            int dx1 = 40000;
            int dy1 = 80000;
            int dx2 = dx1 + 100000;
            int dy2 = dy1 + 100000;

            XSSFClientAnchor colourAnchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, offsetCol, row, offsetCol, row);
            XSSFSimpleShape colourSquare = xssfDrawing.createSimpleShape(colourAnchor);

            colourSquare.setShapeType(ShapeTypes.RECT);
            colourSquare.setLineStyleColor(0, 0, 0);

            if (rgb != null)
            {
                colourSquare.setFillColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
            }

            // Category label, shifted right to clear the colour swatch.
            XSSFClientAnchor labelAnchor = new XSSFClientAnchor(180000, 0, 0, 0, offsetCol, row, offsetCol + 2, row + 1);
            XSSFTextBox labelBox = xssfDrawing.createTextbox(labelAnchor);

            // Replace standard hyphens with non-breaking hyphens to prevent unwanted wrapping.
            String categoryText = point.getCategory().replace('-', '\u2011');
            labelBox.setText(categoryText);

            // Percentage column.
            int percentCol = offsetCol + 2;
            double percent = (total == 0.0 ? 0.0 : point.getValue().doubleValue() / total * 100.0);

            XSSFClientAnchor percentageAnchor = new XSSFClientAnchor(0, 0, 0, 0, percentCol, row, percentCol + 1, row + 1);
            XSSFTextBox percentageBox = xssfDrawing.createTextbox(percentageAnchor);

            percentageBox.setText(String.format(Locale.ROOT, "%.1f%%", percent));
            percentageBox.setTextDirection(TextDirection.HORIZONTAL);
        }

        configureCustomLegendLayout(chart);
    }

    // TODO: Add Javadoc
    private void addDataLabels(XSSFChart chart)
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

    // TODO: Add Javadoc
    private void configureBarChart(XDDFBarChartData barData, XDDFValueAxis valueAxis, ConfiguredChartData config)
    {
        int base = config.getBarScaleLogBase();

        if (base > 1)
        {
            valueAxis.setLogBase(base);
            valueAxis.setMinimum(1.0);
        }

        barData.setBarDirection(config.getBarDirection());
        barData.setBarGrouping(BarGrouping.CLUSTERED);
        barData.setVaryColors(config.getRgbFillBytes() == null);
        barData.setGapWidth(80);
    }

    /**
     * Configures line chart properties such as line smoothing and markers.
     *
     * @param lineData
     *        the line chart data object
     * @param config
     *        the chart configuration
     */
    private void configureLineChart(XDDFLineChartData lineData, ConfiguredChartData config)
    {
        lineData.setVaryColors(false);

        // Iterate through the series using index-based access
        int seriesCount = lineData.getSeriesCount();

        for (int i = 0; i < seriesCount; i++)
        {
            XDDFChartData.Series series = lineData.getSeries(i);

            if (series instanceof XDDFLineChartData.Series)
            {
                XDDFLineChartData.Series lineSeries = (XDDFLineChartData.Series) series;

                lineSeries.setSmooth(false);
                lineSeries.setMarkerStyle(MarkerStyle.CIRCLE);
                lineSeries.setMarkerSize((short) 5);
            }
        }
    }

    /**
     * Configures Scatter/XY chart properties.
     *
     * @param scatterData
     *        the scatter chart data object
     * @param config
     *        the chart configuration
     */
    private void configureScatterChart(XDDFScatterChartData scatterData, ConfiguredChartData config)
    {
        int seriesCount = scatterData.getSeriesCount();

        scatterData.setVaryColors(false);
        scatterData.setStyle(ScatterStyle.SMOOTH_MARKER);

        for (int i = 0; i < seriesCount; i++)
        {
            XDDFScatterChartData.Series scatterSeries = (XDDFScatterChartData.Series) scatterData.getSeries(i);

            scatterSeries.setSmooth(true);
            scatterSeries.setMarkerStyle(MarkerStyle.DIAMOND);
            scatterSeries.setMarkerSize((short) 5);

            XDDFLineProperties lineProperties = new XDDFLineProperties();
            lineProperties.setFillProperties(new XDDFNoFillProperties());
            // lineProperties.setFillProperties(new XDDFSolidFillProperties(XDDFColor.from(new byte[]{52, 120, (byte) 234})));
            // lineProperties.setWidth(1.5);

            scatterSeries.setLineProperties(lineProperties);
        }
    }

    /*
     * Manual plot area layout values obtained from the generated chart XML.
     * These values reserve space on the right for the custom legend.
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
    private void configureCustomLegendLayout(XSSFChart chart)
    {
        CTPlotArea plotArea = chart.getCTChart().getPlotArea();
        CTLayout layout = plotArea.isSetLayout() ? plotArea.getLayout() : plotArea.addNewLayout();

        if (layout.isSetManualLayout())
        {
            layout.unsetManualLayout();
        }

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