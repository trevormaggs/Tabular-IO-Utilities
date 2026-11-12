package io.style;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import util.HexByteConversion;

/**
 * Immutable configuration describing the data, appearance and layout of a chart.
 *
 * <p>
 * A chart configuration can contain one or more chart series. Each series contains category/value
 * pairs and may optionally define individual point colours or a series colour.
 * </p>
 *
 * <p>
 * Instances are created through the nested {@link Builder} class.
 * </p>
 */
public final class ConfiguredChartData
{
    private final ChartTypes type;
    private final String title;
    private final int titleFontPoint;
    private final byte[] titleBytesRGB;
    private final ChartSeries[] series;
    private final int fromCol;
    private final int fromRow;
    private final int toCol;
    private final int toRow;
    private final BarDirection barDirection;
    private final int barScaleLogBase;
    private final byte[] barFillBytesRGB;

    /**
     * Constructs an immutable chart configuration from the specified builder.
     *
     * @param builder
     *        the builder containing the chart configuration properties
     */
    private ConfiguredChartData(Builder builder)
    {
        type = builder.bd_type;
        title = builder.bd_title;
        titleFontPoint = builder.bd_titleFontPoint;
        titleBytesRGB = (builder.bd_rgbTitleBytes == null ? null : builder.bd_rgbTitleBytes.clone());

        series = new ChartSeries[builder.bd_series.size()];

        for (int i = 0; i < series.length; i++)
        {
            series[i] = new ChartSeries(builder.bd_series.get(i));
        }

        fromCol = builder.bd_fromCol;
        fromRow = builder.bd_fromRow;
        toCol = builder.bd_toCol;
        toRow = builder.bd_toRow;
        barDirection = builder.bd_barDirection;
        barScaleLogBase = builder.bd_barScaleLogBase;
        barFillBytesRGB = (builder.bd_barFillBytesRGB == null ? null : builder.bd_barFillBytesRGB.clone());
    }

    private static byte[][] cloneColours(byte[][] colours)
    {
        if (colours == null)
        {
            return null;
        }

        byte[][] copy = new byte[colours.length][];

        for (int i = 0; i < colours.length; i++)
        {
            copy[i] = (colours[i] == null ? null : colours[i].clone());
        }

        return copy;
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
     * @return a copy of the RGB colour bytes, or {@code null} if not specified
     */
    public byte[] getRgbTitleBytes()
    {
        return (titleBytesRGB == null ? null : titleBytesRGB.clone());
    }

    /**
     * Returns the configured chart series.
     *
     * @return a defensive copy of the chart series array
     */
    public ChartSeries[] getChartSeries()
    {
        ChartSeries[] copy = new ChartSeries[series.length];

        for (int i = 0; i < series.length; i++)
        {
            copy[i] = new ChartSeries(series[i]);
        }

        return copy;
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
     * @return a copy of the RGB colour bytes, or {@code null} if not specified
     */
    public byte[] getRgbFillBytes()
    {
        return (barFillBytesRGB == null ? null : barFillBytesRGB.clone());
    }

    /**
     * Represents one chart series, including its legend settings, category/value pairs and optional
     * point-level or series-level colours.
     */
    public static final class ChartSeries
    {
        private final String legendLabel;
        private final int legendFontPoint;
        private final String[] categories;
        private final Number[] values;
        private final byte[][] pointColours;
        private final byte[] seriesColour;

        private ChartSeries(Builder.SeriesBuilder builder)
        {
            legendLabel = builder.label;
            legendFontPoint = builder.legendFontPoint;
            categories = builder.categories.toArray(new String[0]);
            values = builder.values.toArray(new Number[0]);
            pointColours = new byte[builder.pointColours.size()][];

            for (int i = 0; i < pointColours.length; i++)
            {
                pointColours[i] = (builder.pointColours.get(i) == null ? null : builder.pointColours.get(i).clone());
            }

            seriesColour = (builder.seriesColour == null ? null : builder.seriesColour.clone());

        }

        private ChartSeries(ChartSeries source)
        {
            legendLabel = source.legendLabel;
            legendFontPoint = source.legendFontPoint;
            categories = source.categories.clone();
            values = source.values.clone();
            pointColours = cloneColours(source.pointColours);
            seriesColour = (source.seriesColour == null ? null : source.seriesColour.clone());

        }

        /**
         * Returns the legend label describing this series.
         *
         * @return the legend label, or {@code null} if not specified
         */
        public String getLegendLabel()
        {
            return legendLabel;
        }

        /**
         * Returns the font size configured for the chart legend, in points.
         *
         * @return the legend font size in points
         */
        public int getLegendFontPoint()
        {
            return legendFontPoint;
        }

        /**
         * Returns the categories belonging to this series.
         *
         * @return a copy of the category array
         */
        public String[] getCategories()
        {
            return categories.clone();
        }

        /**
         * Returns the numerical values belonging to this series.
         *
         * @return a copy of the value array
         */
        public Number[] getValues()
        {
            return values.clone();
        }

        /**
         * Returns the optional RGB colours configured for individual data points.
         *
         * @return a deep copy of the point colour array; individual entries may be {@code null}
         */
        public byte[][] getPointColours()
        {
            return cloneColours(pointColours);
        }

        /**
         * Returns the optional RGB colour configured for the entire series.
         *
         * @return a copy of the RGB colour bytes, or {@code null} if not specified
         */
        public byte[] getSeriesColour()
        {
            return (seriesColour == null ? null : seriesColour.clone());
        }

        /**
         * Returns the number of data points in this series.
         *
         * @return the number of categories and values in the series
         */
        public int size()
        {
            return categories.length;
        }
    }

    /**
     * Builder for constructing {@link ConfiguredChartData} instances.
     *
     * <p>
     * The chart type is mandatory. Chart-level properties and one or more chart series may then
     * be configured through the fluent builder methods.
     * </p>
     */
    public static final class Builder
    {
        private final ChartTypes bd_type;
        private String bd_title;
        private int bd_titleFontPoint;
        private byte[] bd_rgbTitleBytes;
        private final List<SeriesBuilder> bd_series;
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
         * @throws NullPointerException
         *         if {@code type} is {@code null}
         */
        public Builder(ChartTypes type)
        {
            bd_type = Objects.requireNonNull(type, "Chart type cannot be null");
            bd_titleFontPoint = 12;
            bd_rgbTitleBytes = new byte[]{0, 0, 0};
            bd_series = new ArrayList<>();
            bd_fromCol = 0;
            bd_fromRow = 0;
            bd_toCol = 0;
            bd_toRow = 0;
            bd_barDirection = BarDirection.COL;
            bd_barScaleLogBase = 0;
        }

        /**
         * Sets the chart title.
         *
         * @param title
         *        the chart title; leading and trailing whitespace is removed
         * @return this builder instance for method chaining
         */
        public Builder title(String title)
        {
            bd_title = (title == null ? null : title.trim());
            return this;
        }

        /**
         * Sets the font size of the chart title, in points.
         *
         * @param point
         *        the title font size; negative values are ignored
         * @return this builder instance for method chaining
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
         * Sets the chart title colour using an RGB hexadecimal string, for example
         * {@code "#ED7D31"} for orange.
         *
         * @param hex
         *        the RGB hexadecimal colour string
         * @return this builder instance for method chaining
         * @throws IllegalArgumentException
         *         if the decoded colour is not exactly three bytes
         */
        public Builder titleColourHex(String hex)
        {
            byte[] rgb = HexByteConversion.decode(hex);

            if (rgb == null || rgb.length != 3)
            {
                throw new IllegalArgumentException("Invalid RGB colour [" + hex + "]");
            }

            bd_rgbTitleBytes = rgb;
            return this;
        }

        /**
         * Sets the legend label for the active series.
         *
         * @param label
         *        the legend label; leading and trailing whitespace is removed
         * @return this builder instance for method chaining
         */
        public Builder legendLabel(String label)
        {
            getActiveSeries().label = (label == null ? null : label.trim());
            return this;
        }

        /**
         * Sets the font size configured for the active series legend, in points.
         *
         * @param point
         *        the legend font size; must not be negative
         * @return this builder instance for method chaining
         * @throws IllegalArgumentException
         *         if {@code point} is negative
         */
        public Builder legendFontPoint(int point)
        {
            if (point < 0)
            {
                throw new IllegalArgumentException("Legend font size cannot be negative");
            }

            getActiveSeries().legendFontPoint = point;
            return this;
        }

        /**
         * Adds a category, numerical value and optional RGB colour to the active series.
         *
         * @param category
         *        the category name
         * @param value
         *        the numerical value to be plotted
         * @param rgb
         *        the optional RGB colour represented by exactly three bytes
         * @return this builder instance for method chaining
         */
        public Builder series(String category, Number value, byte[] rgb)
        {
            getActiveSeries().point(category, value, rgb);
            return this;
        }

        /**
         * Adds a category and numerical value to the active series, using an RGB hexadecimal
         * colour string for the data point.
         *
         * @param category
         *        the category name
         * @param value
         *        the numerical value to be plotted
         * @param hexColour
         *        the RGB hexadecimal colour string
         * @return this builder instance for method chaining
         */
        public Builder series(String category, Number value, String hexColour)
        {
            return series(category, value, HexByteConversion.decode(hexColour));
        }

        /**
         * Adds a category and numerical value to the active series without specifying a point
         * colour.
         *
         * @param category
         *        the category name
         * @param value
         *        the numerical value to be plotted
         * @return this builder instance for method chaining
         */
        public Builder series(String category, Number value)
        {
            return series(category, value, (byte[]) null);
        }

        /**
         * Starts a new chart series with the specified legend label and font size.
         *
         * @param label
         *        the legend label for the new series; may be {@code null}
         * @param legendFontPoint
         *        the legend font size in points; must not be negative
         * @return this builder instance for method chaining
         * @throws IllegalArgumentException
         *         if {@code legendFontPoint} is negative
         */
        public Builder addSeries(String label, int legendFontPoint)
        {
            if (legendFontPoint < 0)
            {
                throw new IllegalArgumentException("Legend font size cannot be negative");
            }

            bd_series.add(new SeriesBuilder(label, legendFontPoint));
            return this;
        }

        /**
         * Starts a new chart series using the default legend font size of 12 points.
         *
         * @param label
         *        the legend label for the new series; may be {@code null}
         * @return this builder instance for method chaining
         */
        public Builder addSeries(String label)
        {
            return addSeries(label, 12);
        }

        /**
         * Adds a category, numerical value and optional RGB colour to the active series.
         *
         * @param category
         *        the category name
         * @param value
         *        the numerical value to be plotted
         * @param rgb
         *        the optional RGB colour represented by exactly three bytes
         * @return this builder instance for method chaining
         */
        public Builder point(String category, Number value, byte[] rgb)
        {
            getActiveSeries().point(category, value, rgb);
            return this;
        }

        /**
         * Adds a category and numerical value to the active series, using an RGB hexadecimal
         * colour string for the data point.
         *
         * @param category
         *        the category name
         * @param value
         *        the numerical value to be plotted
         * @param hexColour
         *        the RGB hexadecimal colour string
         * @return this builder instance for method chaining
         */
        public Builder point(String category, Number value, String hexColour)
        {
            return point(category, value, HexByteConversion.decode(hexColour));
        }

        /**
         * Adds a category and numerical value to the active series without specifying a point
         * colour.
         *
         * @param category
         *        the category name
         * @param value
         *        the numerical value to be plotted
         * @return this builder instance for method chaining
         */
        public Builder point(String category, Number value)
        {
            return point(category, value, (byte[]) null);
        }

        /**
         * Sets the RGB colour for the entire active series.
         *
         * @param hex
         *        the RGB hexadecimal colour string
         * @return this builder instance for method chaining
         * @throws IllegalArgumentException
         *         if the decoded colour is not exactly three bytes
         */
        public Builder seriesColourHex(String hex)
        {

            byte[] rgb = HexByteConversion.decode(hex);

            if (rgb == null || rgb.length != 3)
            {
                throw new IllegalArgumentException("Invalid RGB colour: " + hex);
            }

            getActiveSeries().seriesColour = rgb;
            return this;
        }

        /**
         * Sets the starting column index for chart positioning.
         *
         * @param col
         *        the zero-based column index of the chart's top-left corner
         * @return this builder instance for method chaining
         */
        public Builder fromCol(int col)
        {
            bd_fromCol = col;
            return this;
        }

        /**
         * Sets the starting row index for chart positioning.
         *
         * @param row
         *        the zero-based row index of the chart's top-left corner
         * @return this builder instance for method chaining
         */
        public Builder fromRow(int row)
        {

            bd_fromRow = row;
            return this;
        }

        /**
         * Sets the ending column index for chart positioning.
         *
         * @param col
         *        the zero-based column index of the chart's bottom-right corner
         * @return this builder instance for method chaining
         */
        public Builder toCol(int col)
        {
            bd_toCol = col;
            return this;
        }

        /**
         * Sets the ending row index for chart positioning.
         *
         * @param row
         *        the zero-based row index of the chart's bottom-right corner
         * @return this builder instance for method chaining
         */
        public Builder toRow(int row)
        {
            bd_toRow = row;
            return this;
        }

        /**
         * Sets the direction of the bars in the chart.
         *
         * @param direction
         *        the bar direction, {@code "COL"} for vertical columns or {@code "BAR"} for
         *        horizontal bars
         * @return this builder instance for method chaining
         * @throws NullPointerException
         *         if {@code direction} is {@code null}
         * @throws IllegalArgumentException
         *         if {@code direction} is not a valid {@link BarDirection} name
         */
        public Builder barDirection(String direction)
        {
            Objects.requireNonNull(direction, "Bar direction cannot be null");
            bd_barDirection = BarDirection.valueOf(direction.trim().toUpperCase());
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
         * Sets the chart bar fill colour using an RGB hexadecimal string, for example
         * {@code "#0044FF"} for blue.
         *
         * @param hex
         *        the RGB hexadecimal colour string
         * @return this builder instance for method chaining
         * @throws IllegalArgumentException
         *         if the decoded colour is not exactly three bytes
         */
        public Builder fillColourHex(String hex)
        {
            byte[] rgb = HexByteConversion.decode(hex);

            if (rgb == null || rgb.length != 3)
            {
                throw new IllegalArgumentException("Invalid RGB colour: " + hex);
            }

            bd_barFillBytesRGB = rgb;
            return this;
        }

        /**
         * Constructs a {@link ConfiguredChartData} instance using the properties supplied to this
         * builder.
         *
         * @return the constructed chart configuration
         */
        public ConfiguredChartData build()
        {
            return new ConfiguredChartData(this);
        }

        private SeriesBuilder getActiveSeries()
        {
            if (bd_series.isEmpty())
            {
                bd_series.add(new SeriesBuilder(null, 12));
            }

            return bd_series.get(bd_series.size() - 1);
        }

        private static final class SeriesBuilder
        {
            private String label;
            private int legendFontPoint;
            private final List<String> categories;
            private final List<Number> values;
            private final List<byte[]> pointColours;
            private byte[] seriesColour;

            private SeriesBuilder(String label, int legendFontPoint)
            {
                this.label = (label == null ? null : label.trim());
                this.legendFontPoint = legendFontPoint;
                this.categories = new ArrayList<>();
                this.values = new ArrayList<>();
                this.pointColours = new ArrayList<>();
            }

            private void point(String category, Number value, byte[] rgb)
            {
                if (category == null || category.trim().isEmpty())
                {
                    throw new IllegalArgumentException("Category cannot be null or blank");
                }

                Objects.requireNonNull(value, "Value cannot be null");

                if (rgb != null && rgb.length != 3)
                {
                    throw new IllegalArgumentException("RGB colour must contain exactly three bytes");
                }

                categories.add(category);
                values.add(value);
                pointColours.add(rgb == null ? null : rgb.clone());
            }
        }
    }
}