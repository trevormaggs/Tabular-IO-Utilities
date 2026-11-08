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
 * A chart configuration contains chart-level properties and one or more data series. Each series
 * consists of a legend entry, category labels, and corresponding numerical values. The
 * configuration also specifies the chart's position within the worksheet and optional axis scaling.
 * 
 * Instances are constructed using the nested {@link Builder} class. Mutable input arrays are
 * defensively copied to preserve the immutability of the configuration.
 */
public final class ConfiguredChartData
{
    private final ChartTypes type;
    private final String chartTitle;
    private final int chartTitleFontPoint;
    private final byte[] rgbTitleBytes;
    private final int legendFontPoint;
    private final int barScaleLogBase;
    private final int fromCol;
    private final int fromRow;
    private final int toCol;
    private final int toRow;
    private final BarDirection barDirection;
    private final byte[] rgbChartFillBytes;
    private final String seriesName;
    private final List<ChartDataPoint> dataPoints;

    /**
     * Constructs a fully initialised chart configuration from the specified builder.
     *
     * The supplied series list is defensively copied and exposed as an unmodifiable list.
     * Individual series maintain their own defensive copies of category and value arrays.
     *
     * @param builder
     *        the builder containing the chart configuration properties
     *
     * @throws NullPointerException
     *         if {@code builder} or its chart type is {@code null}
     */
    private ConfiguredChartData(Builder builder)
    {
        Objects.requireNonNull(builder, "Builder must not be null");

        this.type = Objects.requireNonNull(builder.bd_type, "Chart type cannot be null");
        this.chartTitle = builder.bd_title;
        this.chartTitleFontPoint = builder.bd_titleFontPoint;
        this.rgbTitleBytes = (builder.bd_rgbTitleBytes != null ? builder.bd_rgbTitleBytes.clone() : null);
        this.legendFontPoint = builder.bd_legendFontPoint;
        this.barScaleLogBase = builder.bd_barScaleLogBase;
        this.fromCol = builder.bd_fromCol;
        this.fromRow = builder.bd_fromRow;
        this.toCol = builder.bd_toCol;
        this.toRow = builder.bd_toRow;
        this.barDirection = builder.bd_barDirection;
        this.rgbChartFillBytes = (builder.bd_rgbFillBytes != null ? builder.bd_rgbFillBytes.clone() : null);
        this.dataPoints = Collections.unmodifiableList(new ArrayList<>(builder.bd_dataPoints));
        this.seriesName = builder.bd_seriesName;
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
        return chartTitle;
    }

    /**
     * Returns the chart title font size in points.
     *
     * @return the font size in points, or zero to retain the default size
     */
    public int getChartTitleFontPoint()
    {
        return chartTitleFontPoint;
    }

    /**
     * Returns the RGB colour bytes used for the chart title.
     *
     * @return the RGB colour bytes, or {@code null} if unspecified
     */
    public byte[] getRgbTitleBytes()
    {
        return (rgbTitleBytes != null ? rgbTitleBytes.clone() : null);
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
     * Returns the logarithmic base configured for the bar value axis.
     *
     * @return the logarithmic base, or zero if logarithmic scaling is not configured
     */
    public int getBarScaleLogBase()
    {
        return barScaleLogBase;
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
     * Returns the RGB colour bytes used for the chart fill.
     *
     * @return the RGB colour bytes, or {@code null} if unspecified
     */
    public byte[] getRgbFillBytes()
    {
        return (rgbChartFillBytes != null ? rgbChartFillBytes.clone() : null);
    }

    public List<ChartDataPoint> getDataPoints()
    {
        return dataPoints;
    }

    public String getSeriesName()
    {
        return seriesName;
    }

    /**
     * Builder for constructing {@link ConfiguredChartData} instances.
     * 
     * The chart type is mandatory. Other chart properties may be specified as required. Unless
     * explicitly configured, the title colour defaults to black and all chart positioning
     * coordinates default to zero.
     */
    public static final class Builder
    {
        private final ChartTypes bd_type;
        private final List<ChartDataPoint> bd_dataPoints;
        private String bd_seriesName;
        private String bd_title;
        private int bd_titleFontPoint;
        private byte[] bd_rgbTitleBytes;
        private int bd_legendFontPoint;
        private int bd_barScaleLogBase;
        private int bd_fromCol;
        private int bd_fromRow;
        private int bd_toCol;
        private int bd_toRow;
        private BarDirection bd_barDirection;
        private byte[] bd_rgbFillBytes;

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
            this.bd_titleFontPoint = 0;
            this.bd_rgbTitleBytes = new byte[]{(byte) 0, (byte) 0, (byte) 0};
            this.bd_fromCol = 0;
            this.bd_fromRow = 0;
            this.bd_toCol = 0;
            this.bd_toRow = 0;
            this.bd_barScaleLogBase = 0;
            this.bd_barDirection = BarDirection.COL;
            this.bd_rgbFillBytes = new byte[]{(byte) 0, (byte) 0, (byte) 0};
            this.bd_legendFontPoint = 12;
            this.bd_seriesName = "";
            this.bd_dataPoints = new ArrayList<>();
        }

        /**
         * Adds a chart data point comprising a category, numerical value and RGB colour. The
         * category identifies the data point in the chart legend.
         *
         * @param cat
         *        the category name. For bar charts, it is displayed along the category axis
         * @param val
         *        the numerical value to be plotted
         * @param hexColour
         *        the RGB hexadecimal colour, for example {@code "#0044FF"} for blue
         * @return this builder instance for method chaining
         */
        public Builder series(String cat, Number val, String hexColour)
        {
            return series(cat, val, HexByteConversion.decode(hexColour));
        }

        /**
         * Adds a chart data point comprising a category, numerical value and RGB colour. The
         * category identifies the data point in the chart legend.
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
            bd_dataPoints.add(new ChartDataPoint(cat, val, rgb));
            return this;
        }

        /**
         * Adds a chart data point comprising a category and numerical value. The category
         * identifies the data point in the chart legend.
         *
         * If no colour is specified, the data point inherits the chart fill colour.
         * 
         * @param cat
         *        the category name. For bar charts, it is displayed along the category axis
         * @param val
         *        the numerical value to be plotted
         * @return this builder instance for method chaining
         */
        public Builder series(String cat, Number val)
        {
            bd_dataPoints.add(new ChartDataPoint(cat, val, null));
            return this;
        }

        public Builder seriesName(String s)
        {
            if (s != null)
            {
                bd_seriesName = s;
            }

            return this;
        }

        /**
         * Sets the chart title.
         *
         * @param t
         *        the chart title; leading and trailing whitespace is removed
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
         * Sets the chart title colour using an RGB hexadecimal string, for example:
         * {@code "#ED7D31"} denotes oranges.
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
                bd_rgbTitleBytes = HexByteConversion.decode(h.trim());
            }

            return this;
        }

        /**
         * Returns the font point or size of the chart title.
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
         * Returns the font point or size of the chart title.
         *
         * @param legendFontPoint
         *        the font point or size of the chart legend
         * @return this builder instance for method chaining
         */
        public Builder legendFontPoint(int point)
        {
            if (point >= 0)
            {
                bd_legendFontPoint = point;
            }

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
         *        the bar direction: {@code "COL"} for vertical columns or
         *        {@code "BAR"} for horizontal bars
         *
         * @throws NullPointerException
         *         if the specified direction is {@code null}
         * @throws IllegalArgumentException
         *         if the specified direction is neither {@code "COL"} nor {@code "BAR"}
         */
        public Builder barDirection(String bd)
        {
            Objects.requireNonNull(bd, "Bar direction must not be null");
            bd_barDirection = BarDirection.valueOf(bd.trim().toUpperCase());

            return this;
        }

        /**
         * Sets the chart fill colour based on the specified RGB hexadecimal string, for example:
         * {@code "#0044FF"} represents blue.
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
                bd_rgbFillBytes = HexByteConversion.decode(h.trim());
            }

            return this;
        }

        /**
         * Constructs a {@link ConfiguredChartData} instance using the properties supplied
         * to this builder.
         *
         * @return the constructed chart configuration
         */
        public ConfiguredChartData build()
        {
            return new ConfiguredChartData(this);
        }
    }

    /**
     * Represents an individual data series within a chart.
     * 
     * A series consists of a legend entry, category labels, and corresponding numerical values.
     * Category and value arrays are defensively copied when the series is constructed and when they
     * are retrieved.
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

            this.category = cat;
            this.value = val;
            this.rgb = (colours == null ? null : colours.clone());
        }

        /**
         * Returns the category associated with this data point.
         *
         * @return the category name, or {@code null} if none was specified
         */
        public String getCategory()
        {
            return category;
        }

        /**
         * Returns the numerical value associated with this data point.
         *
         * @return the value, or {@code null} if none was specified
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