package io.style;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import util.HexByteConversion;

/**
 * Immutable configuration describing the data, appearance, and layout of a chart to be rendered
 * within an Excel worksheet.
 * 
 * A chart configuration contains chart-level properties and a collection of data points. Each data
 * point consists of a category, a numerical value, and an optional RGB colour. The configuration
 * also specifies the series label, chart position, and optional axis scaling.
 * 
 * Instances are constructed using the nested {@link Builder} class. Mutable input arrays are
 * defensively copied to preserve the immutability of the configuration.
 */
public final class ConfiguredChartDataOld
{
    private final ChartTypes type;
    private final String title;
    private final int titleFontPoint;
    private final byte[] titleBytesRGB;
    private final String legendLabel;
    private final int legendFontPoint;
    private final List<ChartDataPoint> series;
    private final int fromCol;
    private final int fromRow;
    private final int toCol;
    private final int toRow;
    private final BarDirection barDirection;
    private final int barScaleLogBase;
    private final byte[] barFillBytesRGB;

    /**
     * Constructs a fully initialised immutable chart configuration from the specified builder.
     *
     * @param builder
     *        the builder containing the chart configuration properties
     *
     * @throws NullPointerException
     *         if {@code builder} or its chart type is {@code null}
     */
    private ConfiguredChartDataOld(Builder builder)
    {
        Objects.requireNonNull(builder, "Builder must not be null");

        this.type = Objects.requireNonNull(builder.bd_type, "Chart type cannot be null");
        this.title = builder.bd_title;
        this.titleFontPoint = builder.bd_titleFontPoint;
        this.titleBytesRGB = (builder.bd_rgbTitleBytes != null ? builder.bd_rgbTitleBytes.clone() : null);
        this.legendLabel = builder.bd_legendLabel;
        this.legendFontPoint = builder.bd_legendFontPoint;
        this.series = Collections.unmodifiableList(new ArrayList<>(builder.bd_series));
        this.fromCol = builder.bd_fromCol;
        this.fromRow = builder.bd_fromRow;
        this.toCol = builder.bd_toCol;
        this.toRow = builder.bd_toRow;
        this.barDirection = builder.bd_barDirection;
        this.barScaleLogBase = builder.bd_barScaleLogBase;
        this.barFillBytesRGB = (builder.bd_barFillBytesRGB != null ? builder.bd_barFillBytesRGB.clone() : null);
    }

    /**
     * Returns the Apache POI chart type.
     *
     * @return the chart type
     */
    public ChartTypes getChartType()
    {
        return type;
    }

    /**
     * Returns the chart title.
     *
     * @return the chart title, or {@code null} if unspecified
     */
    public String getChartTitle()
    {
        return title;
    }

    /**
     * Returns the title font size in points.
     *
     * @return the font size in points, or zero to retain the default size
     */
    public int getChartTitleFontPoint()
    {
        return titleFontPoint;
    }

    /**
     * Returns the RGB colour bytes used for the chart title.
     *
     * @return the RGB colour bytes, or {@code null} if not specified
     */
    public byte[] getRgbTitleBytes()
    {
        return (titleBytesRGB != null ? titleBytesRGB.clone() : null);
    }

    /**
     * Returns the legend label describing the series
     *
     * @return the legend label , or {@code null} if not specified
     */
    public String getLegendLabel()
    {
        return legendLabel;
    }

    /**
     * Returns the font size of the chart legend, in points.
     *
     * @return the legend font size
     */
    public int getLegendFontPoint()
    {
        return legendFontPoint;
    }

    /**
     * Returns the data points belonging to the chart series.
     *
     * @return an unmodifiable list of chart data points
     */
    public List<ChartDataPoint> getSeries()
    {
        return series;
    }

    /**
     * Returns the starting column index of the chart.
     *
     * @return the zero-based starting column index
     */
    public int getFromCol()
    {
        return fromCol;
    }

    /**
     * Returns the starting row index of the chart.
     *
     * @return the zero-based starting row index
     */
    public int getFromRow()
    {
        return fromRow;
    }

    /**
     * Returns the ending column index of the chart.
     *
     * @return the zero-based ending column index
     */
    public int getToCol()
    {
        return toCol;
    }

    /**
     * Returns the ending row index of the chart.
     *
     * @return the zero-based ending row index
     */
    public int getToRow()
    {
        return toRow;
    }

    /**
     * Returns the direction of the bars in the chart.
     *
     * @return the bar direction ({@link BarDirection#COL} for vertical columns or
     *         {@link BarDirection#BAR} for horizontal bars)
     */
    public BarDirection getBarDirection()
    {
        return barDirection;
    }

    /**
     * Returns the logarithmic base configured for the bar value axis.
     *
     * @return the logarithmic base, or zero if logarithmic scaling is not configured
     */
    public int getBarScaleLogBase()
    {
        return barScaleLogBase;
    }

    /**
     * Returns the RGB colour bytes used for the chart bar fill.
     *
     * @return the RGB colour bytes, or {@code null} if not specified
     */
    public byte[] getRgbFillBytes()
    {
        return (barFillBytesRGB != null ? barFillBytesRGB.clone() : null);
    }

    /**
     * Builder for constructing {@link ConfiguredChartDataOld} instances.
     * 
     * The chart type is mandatory. Other chart properties may be specified as required. Unless
     * explicitly configured, the title colour defaults to black and all chart positioning
     * coordinates default to zero.
     */
    public static final class Builder
    {
        private final ChartTypes bd_type;
        private String bd_title;
        private int bd_titleFontPoint;
        private byte[] bd_rgbTitleBytes;
        private String bd_legendLabel;
        private int bd_legendFontPoint;
        private final List<ChartDataPoint> bd_series;
        private int bd_fromCol;
        private int bd_fromRow;
        private int bd_toCol;
        private int bd_toRow;
        private BarDirection bd_barDirection;
        private int bd_barScaleLogBase;
        private byte[] bd_barFillBytesRGB;

        /**
         * Creates a builder for the specified chart type.
         *
         * @param type
         *        the Apache POI chart type; must not be {@code null}
         *
         * @throws NullPointerException
         *         if {@code type} is {@code null}
         */
        public Builder(ChartTypes type)
        {
            this.bd_type = Objects.requireNonNull(type, "Chart type must not be null");
            this.bd_titleFontPoint = 12;
            this.bd_rgbTitleBytes = new byte[]{(byte) 0, (byte) 0, (byte) 0};
            this.bd_legendFontPoint = 12;
            this.bd_series = new ArrayList<>();
            this.bd_fromCol = 0;
            this.bd_fromRow = 0;
            this.bd_toCol = 0;
            this.bd_toRow = 0;
            this.bd_barDirection = BarDirection.COL;
            this.bd_barScaleLogBase = 0;
        }

        /**
         * Sets the chart title.
         *
         * @param t
         *        the chart title. Leading and trailing whitespace is removed
         * @return this builder instance for method chaining
         */
        public Builder title(String t)
        {
            if (t != null)
            {
                bd_title = t.trim();
            }

            return this;
        }

        /**
         * Returns the font point, denoting the size of the chart title.
         *
         * @return the chart title font point
         */
        public Builder titleFontPoint(int point)
        {
            if (point >= 0)
            {
                bd_titleFontPoint = point;
            }

            return this;
        }

        /**
         * Sets the chart title colour using an RGB hexadecimal string, for example:
         * {@code "#ED7D31"} for orange.
         *
         * @param h
         *        the RGB hexadecimal colour string
         * @return this builder instance for method chaining
         *
         * @throws IllegalArgumentException
         *         if the colour string has an invalid format or contains non-hexadecimal characters
         */
        public Builder titleColourHex(String h)
        {
            if (h != null)
            {
                bd_rgbTitleBytes = HexByteConversion.decode(h);
                
                if (bd_rgbTitleBytes != null && bd_rgbTitleBytes.length != 3)
                {
                    throw new IllegalArgumentException("RGB colour must contain exactly three bytes");
                }
            }

            return this;
        }

        /**
         * Sets the legend label describing the series.
         *
         * @param l
         *        the legend name. Leading and trailing whitespace is removed
         * @return this builder instance for method chaining
         */
        public Builder legendLabel(String l)
        {
            if (l != null)
            {
                bd_legendLabel = l.trim();
            }

            return this;
        }

        /**
         * Sets the font size of the chart legend, in points.
         *
         * @param p
         *        the legend font size in points
         * @return this builder instance for method chaining
         */
        public Builder legendFontPoint(int p)
        {
            if (p >= 0)
            {
                bd_legendFontPoint = p;
            }

            return this;
        }

        /**
         * Adds a chart data point comprising a category, numerical value and RGB colour. The
         * category identifies the data point in the chart series.
         *
         * @param cat
         *        the category name. For bar charts, it is displayed along the category axis
         * @param val
         *        the numerical value to be plotted
         * @param rgb
         *        the RGB colour represented by an array of three bytes
         * @return this builder instance for method chaining
         */
        public Builder series(String cat, Number val, byte[] rgb)
        {
            bd_series.add(new ChartDataPoint(cat, val, rgb));
            return this;
        }

        /**
         * Adds a chart data point comprising a category, numerical value and RGB colour. The
         * category identifies the data point in the chart series.
         *
         * @param cat
         *        the category name. For bar charts, it is displayed along the category axis
         * @param val
         *        the numerical value to be plotted
         * @param hexColour
         *        the RGB hexadecimal colour in string, for example {@code "#0044FF"} for blue
         * @return this builder instance for method chaining
         */
        public Builder series(String cat, Number val, String hexColour)
        {
            return series(cat, val, HexByteConversion.decode(hexColour));
        }

        /**
         * Adds a chart data point comprising a category and numerical value. The category
         * identifies the data point in the chart series. If no colour is specified, the data point
         * inherits the chart fill colour.
         * 
         * @param cat
         *        the category name. For bar charts, it is displayed along the category axis
         * @param val
         *        the numerical value to be plotted
         * @return this builder instance for method chaining
         */
        public Builder series(String cat, Number val)
        {
            bd_series.add(new ChartDataPoint(cat, val, null));
            return this;
        }

        /**
         * Sets the starting column index for chart positioning.
         *
         * @param fc
         *        the zero-based column index of the chart's top-left corner
         * @return this builder instance for method chaining
         */
        public Builder fromCol(int fc)
        {
            bd_fromCol = fc;
            return this;
        }

        /**
         * Sets the starting row index for chart positioning.
         *
         * @param fr
         *        the zero-based row index of the chart's top-left corner
         * @return this builder instance for method chaining
         */
        public Builder fromRow(int fr)
        {
            bd_fromRow = fr;
            return this;
        }

        /**
         * Sets the ending column index for chart positioning.
         *
         * @param tc
         *        the zero-based column index of the chart's bottom-right corner
         * @return this builder instance for method chaining
         */
        public Builder toCol(int tc)
        {
            bd_toCol = tc;
            return this;
        }

        /**
         * Sets the ending row index for chart positioning.
         *
         * @param tr
         *        the zero-based row index of the chart's bottom-right corner
         * @return this builder instance for method chaining
         */
        public Builder toRow(int tr)
        {
            bd_toRow = tr;
            return this;
        }

        /**
         * Sets the direction of the bars in the chart.
         *
         * @param bd
         *        the bar direction: {@code "COL"} for vertical columns or {@code "BAR"} for
         *        horizontal bars
         * @return this builder instance for method chaining
         * 
         * @throws IllegalArgumentException
         *         if the specified direction is neither {@code "COL"} nor {@code "BAR"}
         */
        public Builder barDirection(String bd)
        {
            bd_barDirection = BarDirection.valueOf(bd.trim().toUpperCase());
            return this;
        }

        /**
         * Sets the logarithmic base for the bar value axis.
         *
         * @param base
         *        the logarithmic base; values greater than one enable logarithmic scaling
         * @return this builder instance for method chaining
         */
        public Builder barScaleLogBase(int base)
        {
            if (base > 1)
            {
                bd_barScaleLogBase = base;
            }

            return this;
        }

        /**
         * Sets the chart bar fill colour based on the specified RGB hexadecimal string, for
         * example: {@code "#0044FF"} represents blue.
         *
         * @param h
         *        the RGB hexadecimal colour string
         * @return this builder instance for method chaining
         *
         * @throws IllegalArgumentException
         *         if the colour string has an invalid format or contains non-hexadecimal characters
         */
        public Builder fillColourHex(String h)
        {
            if (h != null)
            {
                bd_barFillBytesRGB = HexByteConversion.decode(h);
            }

            return this;
        }

        /**
         * Constructs a {@link ConfiguredChartDataOld} instance using the properties supplied
         * to this builder.
         *
         * @return the constructed chart configuration
         */
        public ConfiguredChartDataOld build()
        {
            return new ConfiguredChartDataOld(this);
        }
    }

    /**
     * Represents an individual data point within a chart series.
     *
     * A data point consists of a category, a numerical value, and an optional RGB colour.
     */
    public static final class ChartDataPoint
    {
        private final String category;
        private final Number value;
        private final byte[] rgb;

        ChartDataPoint(String cat, Number val, byte[] colours)
        {
            if (cat == null || cat.trim().isEmpty())
            {
                throw new IllegalArgumentException("Category cannot be null or empty.");
            }

            if (val == null)
            {
                throw new IllegalArgumentException("Value cannot be null.");
            }

            if (colours != null && colours.length != 3)
            {
                throw new IllegalArgumentException("RGB colour must contain exactly three bytes.");
            }

            this.category = cat;
            this.value = val;
            this.rgb = (colours != null ? colours.clone() : null);
        }

        /**
         * Returns the category associated with this data point.
         *
         * @return the category name
         */
        public String getCategory()
        {
            return category;
        }

        /**
         * Returns the numerical value associated with this data point.
         *
         * @return the value
         */
        public Number getValue()
        {
            return value;
        }

        /**
         * Returns the RGB colour associated with this data point.
         *
         * @return a copy of the RGB colour bytes, or {@code null} if no colour was specified
         */
        public byte[] getColourRGB()
        {
            return (rgb != null ? rgb.clone() : null);
        }
    }
}