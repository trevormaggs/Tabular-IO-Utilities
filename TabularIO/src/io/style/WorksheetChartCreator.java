package io.style;

import java.util.Locale;
import java.util.Objects;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import io.TabularDataException;
import io.style.ConfiguredChartData.ChartSeries;

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
    private boolean scatterLineVisible;

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
     * Overrides the default vertical axis position for subsequently created charts.
     *
     * @param axis
     *        the new vertical axis position
     *
     * @throws NullPointerException
     *         if {@code axis} is {@code null}
     */
    public void setScatterLineVisible(boolean line)
    {
        scatterLineVisible = line;
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
        Objects.requireNonNull(config, "Chart configuration cannot be null");

        ChartTypes type = config.getChartType();
        ChartSeries[] chartSeries = config.getChartSeries();
        boolean isPieBased = (type == ChartTypes.PIE || type == ChartTypes.PIE3D || type == ChartTypes.DOUGHNUT);

        if (chartSeries.length == 0)
        {
            throw new IllegalArgumentException("Chart must contain at least one series");
        }

        if (isPieBased && chartSeries.length != 1)
        {
            throw new IllegalArgumentException("Pie-based charts support exactly one series");
        }

        XSSFClientAnchor anchor = xssfDrawing.createAnchor(0, 0, 0, 0, config.getFromCol(), config.getFromRow(), config.getToCol(), config.getToRow());
        XSSFChart chart = xssfDrawing.createChart(anchor);

        if (type == ChartTypes.BAR3D || type == ChartTypes.LINE3D || type == ChartTypes.PIE3D)
        {
            configure3DView(chart);
        }

        XDDFChartData data;
        XDDFCategoryAxis categoryAxis = null;
        XDDFValueAxis valueAxis = null;

        if (isPieBased)
        {
            data = chart.createData(type, null, null);
        }

        else
        {
            categoryAxis = chart.createCategoryAxis(xAxis);
            valueAxis = chart.createValueAxis(yAxis);
            categoryAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
            data = chart.createData(type, categoryAxis, valueAxis);
        }

        for (ChartSeries item : chartSeries)
        {
            String[] categories = item.getCategories();
            Number[] values = item.getValues();
            byte[] seriesColour = item.getSeriesColour();
            XDDFCategoryDataSource categoryData = XDDFDataSourcesFactory.fromArray(categories);
            XDDFNumericalDataSource<? extends Number> valueData = XDDFDataSourcesFactory.fromArray(values);
            XDDFChartData.Series series = data.addSeries(categoryData, valueData);

            if (!isPieBased && item.getLegendLabel() != null)
            {
                series.setTitle(item.getLegendLabel(), null);
            }

            if (seriesColour != null)
            {
                XDDFShapeProperties properties = new XDDFShapeProperties();

                properties.setFillProperties(new XDDFSolidFillProperties(XDDFColor.from(seriesColour)));
                series.setShapeProperties(properties);
            }
        }

        configureChartTitle(chart, config);

        if (data instanceof XDDFBarChartData)
        {
            int base = config.getBarScaleLogBase();

            if (base > 1)
            {
                valueAxis.setLogBase(base);
                valueAxis.setMinimum(1.0);
            }

            configureBarChart(chart, (XDDFBarChartData) data, config);
        }

        else if (data instanceof XDDFLineChartData)
        {
            configureLineChart((XDDFLineChartData) data, config);
        }

        else if (data instanceof XDDFAreaChartData)
        {
            ((XDDFAreaChartData) data).setVaryColors(false);
        }

        else if (data instanceof XDDFDoughnutChartData)
        {
            ((XDDFDoughnutChartData) data).setHoleSize(60);
        }

        else if (data instanceof XDDFScatterChartData)
        {
            configureScatterChart((XDDFScatterChartData) data, config);
        }

        if (!isPieBased)
        {
            configureChartLegend(chart, config);
        }

        chart.plot(data);

        if (isPieBased)
        {
            drawCustomLegendBox(chart, config);
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
        int fontPoint = config.getChartSeries()[0].getLegendFontPoint();
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

        if (fontPoint > 0)
        {
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

                paragraph.getPPr().getDefRPr().setSz(fontPoint * 100);
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
     * This method supports charts containing exactly one series.
     * </p>
     *
     * @param chart
     *        the pie based chart to configure
     * @param config
     *        the chart configuration containing the single series of category, value and colour
     *        data
     */
    private void drawCustomLegendBox(XSSFChart chart, ConfiguredChartData config)
    {
        double total = 0.0;
        CTPieSer[] seriesArray;
        ChartSeries chartSeries = config.getChartSeries()[0];
        String[] categories = chartSeries.getCategories();
        Number[] values = chartSeries.getValues();
        byte[][] pointColours = chartSeries.getPointColours();
        int count = categories.length;

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

        // Calculate the total before determining the percentage of each data point.
        for (int i = 0; i < count; i++)
        {
            double value = values[i].doubleValue();

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

        // Apply configured point colours to the chart slices.
        for (int i = 0; i < count; i++)
        {
            byte[] rgb = pointColours[i];

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

        for (int i = 0; i < count; i++)
        {
            int row = offsetRow + (i * 2);
            byte[] rgb = pointColours[i];

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
            String categoryText = categories[i].replace('-', '\u2011');
            labelBox.setText(categoryText);

            // Percentage column.
            int percentCol = offsetCol + 2;
            double percent = (total == 0.0 ? 0.0 : values[i].doubleValue() / total * 100.0);
            XSSFClientAnchor percentageAnchor = new XSSFClientAnchor(0, 0, 0, 0, percentCol, row, percentCol + 1, row + 1);

            XSSFTextBox percentageBox = xssfDrawing.createTextbox(percentageAnchor);
            percentageBox.setText(String.format(Locale.ROOT, "%.1f%%", percent));
            percentageBox.setTextDirection(TextDirection.HORIZONTAL);
        }

        configureCustomLegendLayout(chart);
    }

    /**
     * Configures the properties of a bar chart.
     *
     * @param barData
     *        the bar chart data object
     * @param valueAxis
     *        the Y axis object
     * @param config
     *        the chart configuration
     */
    private void configureBarChart(XSSFChart chart, XDDFBarChartData barData, ConfiguredChartData config)
    {
        byte[] rgbFillBytes = config.getRgbFillBytes();
        CTBarChart ctBarChart = chart.getCTChart().getPlotArea().getBarChartArray(0);

        barData.setBarDirection(config.getBarDirection());
        barData.setBarGrouping(BarGrouping.CLUSTERED);
        barData.setVaryColors(rgbFillBytes == null);
        barData.setGapWidth(80);

        if (rgbFillBytes != null)
        {
            for (int i = 0; i < barData.getSeriesCount(); i++)
            {
                XDDFChartData.Series series = barData.getSeries(i);
                XDDFShapeProperties shapeProperties = new XDDFShapeProperties();

                shapeProperties.setFillProperties(new XDDFSolidFillProperties(XDDFColor.from(rgbFillBytes)));
                series.setShapeProperties(shapeProperties);
            }
        }

        for (CTBarSer barSeries : ctBarChart.getSerArray())
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
        scatterData.setStyle(scatterLineVisible ? ScatterStyle.SMOOTH_MARKER : ScatterStyle.MARKER);

        for (int i = 0; i < seriesCount; i++)
        {
            XDDFLineProperties lineProperties = new XDDFLineProperties();
            XDDFScatterChartData.Series scatterSeries = (XDDFScatterChartData.Series) scatterData.getSeries(i);

            scatterSeries.setSmooth(scatterLineVisible);
            scatterSeries.setMarkerStyle(MarkerStyle.DIAMOND);
            scatterSeries.setMarkerSize((short) 5);

            if (scatterLineVisible)
            {
                lineProperties.setFillProperties(new XDDFSolidFillProperties(XDDFColor.from(new byte[]{52, 120, (byte) 234})));
                lineProperties.setWidth(1.5);
            }

            else
            {
                lineProperties.setFillProperties(new XDDFNoFillProperties());
            }

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