package io.style;

import java.util.Objects;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrossBetween;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import io.TabularDataException;

/**
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

    public WorksheetChartCreator(XSSFSheet sheet) throws TabularDataException
    {
        this.xssfSheet = Objects.requireNonNull(sheet, "Worksheet must not be null");
        this.xssfDrawing = this.xssfSheet.createDrawingPatriarch();
        this.xAxis = DEFAULT_XAXIS;
        this.yAxis = DEFAULT_YAXIS;
    }

    public XSSFSheet getXssfSheet()
    {
        return xssfSheet;
    }

    public void overrideDefaultXAxis(AxisPosition axis)
    {
        xAxis = Objects.requireNonNull(axis, "X-axis position must not be null");
    }

    public void overrideDefaultYAxis(AxisPosition axis)
    {
        yAxis = Objects.requireNonNull(axis, "Y-axis position must not be null");
    }

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
                addDoughnutDataLabels(chart);
            }
        }
    }

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

    private static void setBarScaleLogBase(XDDFValueAxis valueAxis, int base)
    {
        if (base > 1)
        {
            valueAxis.setLogBase(base);
            valueAxis.setMinimum(1.0);
        }
    }

    private static void configureBarChart(XDDFBarChartData barData, XDDFValueAxis valueAxis, ConfiguredChartData config)
    {
        barData.setBarDirection(config.getBarDirection());
        barData.setBarGrouping(BarGrouping.CLUSTERED);
        barData.setVaryColors(config.getRgbFillBytes() == null);
        barData.setGapWidth(80);

        setBarScaleLogBase(valueAxis, config.getBarScaleLogBase());
    }

    private static void configureDoughnutChart(XDDFDoughnutChartData doughnutData)
    {
        doughnutData.setVaryColors(true);
        doughnutData.setHoleSize(60);
    }

    /**
     * Configures the font size and optional font colour of a chart title.
     *
     * <p>
     * The title formatting properties are obtained from the supplied chart configuration.
     * The font size is applied to the title paragraphs and their text runs. If an RGB
     * colour is supplied, it is applied to the title text.
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

    private void addDoughnutDataLabels(XSSFChart chart)
    {
        CTDoughnutChart doughnutChart =
                chart.getCTChart().getPlotArea().getDoughnutChartArray(0);

        for (CTPieSer series : doughnutChart.getSerArray())
        {
            CTDLbls dataLabels = series.addNewDLbls();

            dataLabels.addNewShowPercent().setVal(true);
            dataLabels.addNewShowCatName().setVal(true);

            series.addNewDLbls().addNewShowLeaderLines().setVal(true);
        }
    }
}