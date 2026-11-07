package io.style;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/**
 * Immutable configuration for a worksheet chart legend.
 *
 * <p>
 * Calculates percentages automatically from the numerical values supplied for each legend entry.
 * </p>
 * 
 * <b>Change log:</b>
 * <ul>
 * <li>7 November 2026 - Initial creation.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.1
 * @since 7 November 2026
 */
public final class ConfiguredLegendData
{
    private final List<LegendEntry> entries;

    private ConfiguredLegendData(Builder builder)
    {
        double total = 0.0;

        for (LegendEntry entry : builder.bd_entries)
        {
            total += entry.value;
        }

        List<LegendEntry> configuredEntries = new ArrayList<LegendEntry>();

        for (LegendEntry entry : builder.bd_entries)
        {
            double percentage = (total == 0.0 ? 0.0 : entry.value / total * 100.0);
            configuredEntries.add(new LegendEntry(entry.label, entry.value, percentage, entry.rgbBytes));
        }

        this.entries = Collections.unmodifiableList(configuredEntries);
    }

    /**
     * Returns the configured legend entries.
     *
     * @return an unmodifiable list of legend entries
     */
    public List<LegendEntry> getEntries()
    {
        return entries;
    }

    /**
     * Represents an individual entry in a worksheet chart legend.
     */
    public static final class LegendEntry
    {
        private final String label;
        private final double value;
        private final double percentage;
        private final byte[] rgbBytes;

        private LegendEntry(String label, double value, double percentage, byte[] rgbBytes)
        {
            this.label = label;
            this.value = value;
            this.percentage = percentage;
            this.rgbBytes = rgbBytes != null ? rgbBytes.clone() : null;
        }

        public String getLabel()
        {
            return label;
        }

        public double getValue()
        {
            return value;
        }

        public double getPercentage()
        {
            return percentage;
        }

        public String getPercentageText()
        {
            return String.format(Locale.ROOT, "%.1f%%", percentage);
        }

        public byte[] getRgbBytes()
        {
            return rgbBytes != null ? rgbBytes.clone() : null;
        }
    }

    /**
     * Builder for {@link ConfiguredLegendData}.
     */
    public static final class Builder
    {
        private final List<LegendEntry> bd_entries = new ArrayList<LegendEntry>();

        /**
         * Adds an entry to the legend configuration.
         *
         * @param label
         *        the legend entry label
         * @param value
         *        the numerical value used to calculate its percentage
         * @param rgbBytes
         *        the RGB colour as three bytes, or null if no colour is specified
         * @return this builder instance
         */
        public Builder entry(String label, double value, byte[] rgbBytes)
        {
            Objects.requireNonNull(label, "Legend label must not be null");

            if (label.trim().isEmpty())
            {
                throw new IllegalArgumentException("Legend label must not be empty");
            }

            if (Double.isNaN(value) || Double.isInfinite(value) || value < 0.0)
            {
                throw new IllegalArgumentException("Legend value must be a finite, non-negative number");
            }

            if (rgbBytes != null && rgbBytes.length != 3)
            {
                throw new IllegalArgumentException("RGB colour must contain exactly three bytes");
            }

            bd_entries.add(new LegendEntry(label.trim(), value, 0.0, rgbBytes));

            return this;
        }

        /**
         * Builds the immutable legend configuration.
         *
         * @return a new {@link ConfiguredLegendData} instance
         */
        public ConfiguredLegendData build()
        {
            return new ConfiguredLegendData(this);
        }
    }
}