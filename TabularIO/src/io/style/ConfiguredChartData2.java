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
public final class ConfiguredChartData2
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

    private ConfiguredChartData2(Builder builder)
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

    public ChartTypes getChartType()
    {
        return type;
    }

    public String getChartTitle()
    {
        return title;
    }

    public int getChartTitleFontPoint()
    {
        return titleFontPoint;
    }

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

    public int getFromCol()
    {
        return fromCol;
    }

    public int getFromRow()
    {
        return fromRow;
    }

    public int getToCol()
    {
        return toCol;
    }

    public int getToRow()
    {
        return toRow;
    }

    public BarDirection getBarDirection()
    {
        return barDirection;
    }

    public int getBarScaleLogBase()
    {
        return barScaleLogBase;
    }

    public byte[] getRgbFillBytes()
    {
        return (barFillBytesRGB == null ? null : barFillBytesRGB.clone());
    }

    /**
     * Represents one chart series.
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

        public String getLegendLabel()
        {
            return legendLabel;
        }

        public int getLegendFontPoint()
        {
            return legendFontPoint;
        }

        public String[] getCategories()
        {
            return categories.clone();
        }

        public Number[] getValues()
        {
            return values.clone();
        }

        public byte[][] getPointColours()
        {
            return cloneColours(pointColours);
        }

        public byte[] getSeriesColour()
        {
            return (seriesColour == null ? null : seriesColour.clone());
        }

        public int size()
        {
            return categories.length;
        }
    }

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

        public Builder title(String title)
        {
            bd_title = (title == null ? null : title.trim());
            return this;
        }

        public Builder titleFontPoint(int point)
        {
            if (point >= 0)
            {
                bd_titleFontPoint = point;
            }

            return this;
        }

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

        public Builder legendLabel(String label)
        {
            getActiveSeries().label = (label == null ? null : label.trim());
            return this;
        }

        public Builder legendFontPoint(int point)
        {
            if (point < 0)
            {
                throw new IllegalArgumentException("Legend font size cannot be negative");
            }

            getActiveSeries().legendFontPoint = point;
            return this;
        }

        public Builder series(String category, Number value, byte[] rgb)
        {
            getActiveSeries().point(category, value, rgb);
            return this;
        }

        public Builder series(String category, Number value, String hexColour)
        {
            return series(category, value, HexByteConversion.decode(hexColour));
        }

        public Builder series(String category, Number value)
        {
            return series(category, value, (byte[]) null);
        }

        public Builder addSeries(String label, int legendFontPoint)
        {
            if (legendFontPoint < 0)
            {
                throw new IllegalArgumentException("Legend font size cannot be negative");
            }

            bd_series.add(new SeriesBuilder(label, legendFontPoint));
            return this;
        }

        public Builder addSeries(String label)
        {
            return addSeries(label, 12);
        }

        public Builder point(String category, Number value, byte[] rgb)
        {
            getActiveSeries().point(category, value, rgb);
            return this;
        }

        public Builder point(String category, Number value, String hexColour)
        {
            return point(category, value, HexByteConversion.decode(hexColour));
        }

        public Builder point(String category, Number value)
        {
            return point(category, value, (byte[]) null);
        }

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

        public Builder fromCol(int col)
        {
            bd_fromCol = col;
            return this;
        }

        public Builder fromRow(int row)
        {
            bd_fromRow = row;
            return this;
        }

        public Builder toCol(int col)
        {
            bd_toCol = col;
            return this;
        }

        public Builder toRow(int row)
        {
            bd_toRow = row;
            return this;
        }

        public Builder barDirection(String direction)
        {
            Objects.requireNonNull(direction, "Bar direction cannot be null");
            bd_barDirection = BarDirection.valueOf(direction.trim().toUpperCase());
            return this;
        }

        public Builder barScaleLogBase(int base)
        {
            if (base > 1)
            {
                bd_barScaleLogBase = base;
            }

            return this;
        }

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

        public ConfiguredChartData2 build()
        {
            return new ConfiguredChartData2(this);
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