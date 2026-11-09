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

        boolean isPieBased = false;
        ChartTypes type = config.getChartType();
        XDDFCategoryAxis categoryAxis = null;
        XDDFValueAxis valueAxis = null;
        XDDFChartData data;
        XSSFClientAnchor anchor = xssfDrawing.createAnchor(0, 0, 0, 0, config.getFromCol(), config.getFromRow(), config.getToCol(), config.getToRow());
        XSSFChart chart = xssfDrawing.createChart(anchor);

        if (config.getChartTitle() != null)
        {
            chart.setTitleText(config.getChartTitle());
            chart.setTitleOverlay(false);
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
                    XDDFView3D view3D = chart.getOrAddView3D();
                    view3D.setXRotationAngle((byte) 30);
                    view3D.setYRotationAngle(0);
                    view3D.setPerspectiveAngle((short) 30);
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

        List<ChartDataPoint> dataPoint = config.getSeries();
        int len = dataPoint.size();
        String[] catArr = new String[len];
        Number[] valArr = new Number[len];

        for (int i = 0; i < len; i++)
        {
            catArr[i] = dataPoint.get(i).getCategory();
            valArr[i] = dataPoint.get(i).getValue();
        }

        XDDFCategoryDataSource categories = XDDFDataSourcesFactory.fromArray(catArr);
        XDDFNumericalDataSource<? extends Number> values = XDDFDataSourcesFactory.fromArray(valArr);
        XDDFChartData.Series series = data.addSeries(categories, values);

        configureChartTitle(chart, config);

        if (!isPieBased)
        {
            configureChartLegend(chart, config);
            series.setTitle(config.getLegendLabel(), null);

            byte[] rgbFillBytes = config.getRgbFillBytes();

            if (rgbFillBytes != null)
            {
                XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(rgbFillBytes));
                XDDFShapeProperties shapeProperties = new XDDFShapeProperties();

                shapeProperties.setFillProperties(fill);
                series.setShapeProperties(shapeProperties);
            }
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
     * Adds a custom legend to a pie based chart.
     *
     * <p>
     * Creates coloured swatches, category labels and calculated percentages using worksheet drawing
     * objects. Also applies the configured colours to the chart slices and adjusts the plot area to
     * accommodate the custom legend.
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
        List<String> categories = new ArrayList<>();
        List<Double> values = new ArrayList<>();
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

        // Remove the native chart legend. A new custom worksheet based legend replaces it instead.
        if (chart.getCTChart().isSetLegend())
        {
            chart.getCTChart().unsetLegend();
        }

        for (CTPieSer series : seriesArray)
        {
            CTNumData valueData = null;
            CTStrData categoryData = null;
            CTNumDataSource valueSource = series.getVal();
            CTAxDataSource categorySource = series.getCat();

            // Remove native data labels. New percentages are displayed in the custom legend box.
            if (series.isSetDLbls())
            {
                series.unsetDLbls();
            }

            // Obtain categories from literal data or a string reference cache.
            if (categorySource.isSetStrLit())
            {
                categoryData = categorySource.getStrLit();
            }

            else if (categorySource.isSetStrRef())
            {
                CTStrRef strRef = categorySource.getStrRef();

                if (strRef.isSetStrCache())
                {
                    categoryData = strRef.getStrCache();
                }
            }

            // Obtain values from literal data or a numeric reference cache.
            if (valueSource.isSetNumLit())
            {
                valueData = valueSource.getNumLit();
            }

            else if (valueSource.isSetNumRef())
            {
                CTNumRef numRef = valueSource.getNumRef();

                if (numRef.isSetNumCache())
                {
                    valueData = numRef.getNumCache();
                }
            }

            if (categoryData != null)
            {
                for (CTStrVal point : categoryData.getPtArray())
                {
                    categories.add(point.getV());
                }
            }

            if (valueData != null)
            {
                for (CTNumVal point : valueData.getPtArray())
                {
                    double value = Double.parseDouble(point.getV());

                    values.add(Double.valueOf(value));
                    total += value;
                }
            }

            if (categories.size() != values.size())
            {
                throw new IllegalStateException("Category count does not match value count.");
            }

            if (categories.size() != dataPoints.size())
            {
                throw new IllegalStateException("Chart category count does not match configured data point count.");
            }

            for (int i = 0; i < categories.size(); i++)
            {
                byte[] rgb = dataPoints.get(i).getColourRGB();

                CTDPt dataPoint = series.addNewDPt();
                dataPoint.addNewIdx().setVal(i);
                dataPoint.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(rgb);
            }
        }

        int offsetCol = config.getFromCol() + 5;
        int offsetRow = config.getFromRow() + 6;

        for (int i = 0; i < categories.size(); i++)
        {
            int row = offsetRow + (i * 2);
            ChartDataPoint point = dataPoints.get(i);
            byte[] rgb = (point.getColourRGB() == null ? config.getRgbFillBytes() : point.getColourRGB());

            // Colour square gets moved to the left side of offsetCol
            int dx1 = 40000;
            int dy1 = 80000;
            int dx2 = dx1 + 100000;
            int dy2 = dy1 + 100000;

            XSSFClientAnchor colourAnchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, offsetCol, row, offsetCol, row);
            XSSFSimpleShape colourSquare = xssfDrawing.createSimpleShape(colourAnchor);

            colourSquare.setShapeType(ShapeTypes.RECT);
            colourSquare.setLineStyleColor(0, 0, 0);
            colourSquare.setFillColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);

            // Category label is shifted right with dx1 = 180000 to clear the swatch
            XSSFClientAnchor labelAnchor = new XSSFClientAnchor(180000, 0, 0, 0, offsetCol, row, offsetCol + 2, row + 1);
            XSSFTextBox labelBox = xssfDrawing.createTextbox(labelAnchor);

            // Removes standard Hyphens "-" to stop it from wrapping over.
            // \u2011 is a non-breaking hyphen unicode character
            String categoryText = categories.get(i).replace('-', '\u2011');
            labelBox.setText(categoryText);

            int percentCol = offsetCol + 2;
            double percent = (total == 0.0 ? 0.0 : values.get(i).doubleValue() / total * 100.0);

            XSSFClientAnchor percentageAnchor = new XSSFClientAnchor(0, 0, 0, 0, percentCol, row, percentCol + 1, row + 1);
            XSSFTextBox percentageBox = xssfDrawing.createTextbox(percentageAnchor);
            percentageBox.setText(String.format(Locale.ROOT, "%.1f%%", percent));
            percentageBox.setTextDirection(TextDirection.HORIZONTAL);
        }

        configureCustomLegendLayout(chart);
    }

    /*
     * Manual plot area layout values obtained from the generated chart XML. These values
     * reserve space on the right for the custom legend. Example from raw chart.xml below:
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
        CTManualLayout manualLayout = layout.addNewManualLayout();

        manualLayout.addNewLayoutTarget().setVal(STLayoutTarget.INNER);
        manualLayout.addNewXMode().setVal(STLayoutMode.EDGE);
        manualLayout.addNewYMode().setVal(STLayoutMode.EDGE);

        manualLayout.addNewX().setVal(0.054404117454068243);
        manualLayout.addNewY().setVal(0.1295334645669291);
        manualLayout.addNewW().setVal(0.49714587434383201);
        manualLayout.addNewH().setVal(0.79543339895013121);
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
     * @param chart
     *        the pie based chart to configure
     * @param config
     *        the chart configuration containing the category, value and colour data points
     */
    @Deprecated
    private void drawCustomLegendBoxTest(XSSFChart chart, ConfiguredChartData config)
    {
        Objects.requireNonNull(chart, "Chart must not be null");
        Objects.requireNonNull(config, "Config must not be null");

        List<ChartDataPoint> dataPoints = config.getSeries();

        if (dataPoints == null || dataPoints.isEmpty())
        {
            return;
        }

        CTPieSer series;

        switch (config.getChartType())
        {
            case DOUGHNUT:
                series = chart.getCTChart().getPlotArea().getDoughnutChartArray(0).getSerArray(0);
            break;

            case PIE:
                series = chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0);
            break;

            case PIE3D:
                series = chart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0);
            break;

            default:
                throw new IllegalArgumentException("Unsupported chart type.");
        }

        // Remove the native chart legend and data labels
        if (chart.getCTChart().isSetLegend())
        {
            chart.getCTChart().unsetLegend();
        }

        if (series.isSetDLbls())
        {
            series.unsetDLbls();
        }

        double total = 0.0;
        for (ChartDataPoint point : dataPoints)
        {
            total += point.getValue().doubleValue();
        }

        // Apply custom colors to series data points and extract data directly from config
        for (int i = 0; i < dataPoints.size(); i++)
        {
            ChartDataPoint point = dataPoints.get(i);
            byte[] rgb = point.getColourRGB();

            if (rgb != null && rgb.length == 3)
            {
                CTDPt dataPoint = series.addNewDPt();
                dataPoint.addNewIdx().setVal(i);
                dataPoint.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(rgb);
            }
        }

        // Configure manual layout to reserve space for the legend
        CTPlotArea plotArea = chart.getCTChart().getPlotArea();
        CTLayout layout = plotArea.isSetLayout() ? plotArea.getLayout() : plotArea.addNewLayout();
        CTManualLayout manualLayout = layout.addNewManualLayout();

        manualLayout.addNewLayoutTarget().setVal(STLayoutTarget.INNER);
        manualLayout.addNewXMode().setVal(STLayoutMode.EDGE);
        manualLayout.addNewYMode().setVal(STLayoutMode.EDGE);
        manualLayout.addNewX().setVal(0.054404117454068243);
        manualLayout.addNewY().setVal(0.1295334645669291);
        manualLayout.addNewW().setVal(0.49714587434383201);
        manualLayout.addNewH().setVal(0.79543339895013121);

        // Render shapes and text boxes on the sheet
        int offsetCol = config.getFromCol() + 5;
        int offsetRow = config.getFromRow() + 6;

        for (int i = 0; i < dataPoints.size(); i++)
        {
            ChartDataPoint point = dataPoints.get(i);
            int row = offsetRow + (i * 2);
            byte[] rgb = point.getColourRGB();

            // Colour square
            XSSFClientAnchor colourAnchor = new XSSFClientAnchor(20000, 30000, 90000, 70000, offsetCol, row, offsetCol, row + 1);
            XSSFSimpleShape colourSquare = xssfDrawing.createSimpleShape(colourAnchor);

            colourSquare.setShapeType(ShapeTypes.RECT);
            colourSquare.setLineStyleColor(0, 0, 0);

            if (rgb != null && rgb.length == 3)
            {
                colourSquare.setFillColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
            }

            // Category label
            XSSFClientAnchor labelAnchor = new XSSFClientAnchor(40000, 0, 0, 0, offsetCol, row, offsetCol + 2, row + 1);
            XSSFTextBox labelBox = xssfDrawing.createTextbox(labelAnchor);
            labelBox.setText(point.getCategory());

            // Percentage column
            int percentCol = offsetCol + 2;
            double percent = (total == 0.0 ? 0.0 : point.getValue().doubleValue() / total * 100.0);

            XSSFClientAnchor percentageAnchor = new XSSFClientAnchor(0, 0, 0, 0, percentCol, row, percentCol + 1, row + 1);
            XSSFTextBox percentageBox = xssfDrawing.createTextbox(percentageAnchor);
            percentageBox.setText(String.format(Locale.ROOT, "%.1f%%", percent));
            percentageBox.setTextDirection(TextDirection.HORIZONTAL);
        }
    }
}