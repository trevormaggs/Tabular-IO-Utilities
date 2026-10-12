package io.style;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Applies Apache POI conditional formatting rules to manage header and alternating-row appearance
 * across specified cell ranges in a worksheet.
 *
 * <p>
 * Six predefined color styles are available: Grey, Coral, Soft Purple, Lime, Orange, and Sky Blue.
 * </p>
 *
 * <p>
 * <b>Usage Examples:</b>
 * </p>
 *
 * <ul>
 * <li><b>Predefined Style:</b>
 * 
 * <pre>{@code
 * AlternatingRowFactory factory = AlternatingRowFactory.newInstance(sheet, ColourStyle.CORAL);
 * factory.apply(0, 0, rowCount - 1, colCount - 1);
 * }</pre>
 * 
 * </li>
 *
 * <li><b>Custom RGB via Fluent Builder:</b>
 * 
 * <pre>{@code
 * AlternatingRowFactory factory = AlternatingRowFactory.newInstance(sheet)
 *         .setHeaderRowRGB(192, 80, 77)
 *         .setEvenRowRGB(204, 255, 204)
 *         .setOddRowRGB(128, 100, 162)
 *         .setBorderRGB(128, 128, 128);
 * factory.apply(0, 0, rowCount - 1, colCount - 1);
 * }</pre>
 * 
 * </li>
 * </ul>
 *
 * <p>
 * <b>Change log:</b>
 * </p>
 * <ul>
 * <li>24 November 2021 - Initial creation.</li>
 * <li>6 October 2026 - Fixed cell range bounding logic and updated documentation to match POI
 * signatures.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.5
 * @since 24 November 2021
 */
public final class AlternatingRowFactory
{
    private static final String EVEN_ROW_RULE = "MOD(ROW(), 2) = 0";
    private static final String ODD_ROW_RULE = "MOD(ROW(), 2) = 1";
    private final Sheet sheet;
    private final SheetConditionalFormatting formatter;
    private Color borderColour;
    private Color oddRowColour;
    private Color evenRowColour;
    private Color headerRowColour;

    /**
     * Private constructor initialising the conditional formatting context for the target worksheet.
     *
     * @param tab
     *        the target {@link Sheet} where styles will be applied
     * @param colour
     *        the initial {@link ColourStyle} configuration
     */
    private AlternatingRowFactory(Sheet tab, ColourStyle colour)
    {
        this.sheet = tab;
        this.formatter = tab.getSheetConditionalFormatting();

        setPredefinedColour(colour);
    }

    /**
     * Sets the header-row background colour using the specified RGB values.
     *
     * @param red
     *        the red component, from 0 to 255
     * @param green
     *        the green component, from 0 to 255
     * @param blue
     *        the blue component, from 0 to 255
     * @return this factory instance
     */
    public AlternatingRowFactory setHeaderRowRGB(int red, int green, int blue)
    {
        validateRGB(red, green, blue);
        headerRowColour = getCustomBackgroundRGB(red, green, blue, HSSFColor.HSSFColorPredefined.RED);
        return this;
    }

    /**
     * Sets the even-row background colour using the specified RGB values.
     *
     * @param red
     *        the red component, from 0 to 255
     * @param green
     *        the green component, from 0 to 255
     * @param blue
     *        the blue component, from 0 to 255
     * @return this factory instance
     */
    public AlternatingRowFactory setEvenRowRGB(int red, int green, int blue)
    {
        validateRGB(red, green, blue);
        evenRowColour = getCustomBackgroundRGB(red, green, blue, HSSFColor.HSSFColorPredefined.BRIGHT_GREEN);
        return this;
    }

    /**
     * Sets the odd-row background colour using the specified RGB values.
     *
     * @param red
     *        the red component, from 0 to 255
     * @param green
     *        the green component, from 0 to 255
     * @param blue
     *        the blue component, from 0 to 255
     * @return this factory instance
     */
    public AlternatingRowFactory setOddRowRGB(int red, int green, int blue)
    {
        validateRGB(red, green, blue);
        oddRowColour = getCustomBackgroundRGB(red, green, blue, HSSFColor.HSSFColorPredefined.GREEN);
        return this;
    }

    /**
     * Sets the border colour using the specified RGB values.
     *
     * @param red
     *        the red component, from 0 to 255
     * @param green
     *        the green component, from 0 to 255
     * @param blue
     *        the blue component, from 0 to 255
     * @return this factory instance
     */
    public AlternatingRowFactory setBorderRGB(int red, int green, int blue)
    {
        validateRGB(red, green, blue);
        borderColour = getCustomBackgroundRGB(red, green, blue);
        return this;
    }

    /**
     * Applies header and alternating-row conditional formatting to the specified cell range.
     * 
     * The first row of the range is formatted as the header row. If the range contains additional
     * rows, they are formatted using separate conditional formatting rules for even and
     * odd numbered worksheet rows.
     * 
     * If the supplied range is empty or reversed, no formatting is applied and this factory
     * instance is returned unchanged.
     * 
     * @param firstRow
     *        the first row index, inclusive
     * @param firstCol
     *        the first column index, inclusive
     * @param lastRow
     *        the last row index, inclusive
     * @param lastCol
     *        the last column index, inclusive
     * @return this factory instance
     */
    public AlternatingRowFactory apply(int firstRow, int firstCol, int lastRow, int lastCol)
    {
        if (lastRow < firstRow || lastCol < firstCol)
        {
            return this;
        }

        CellRangeAddress[] headerRegion = {new CellRangeAddress(firstRow, firstRow, firstCol, lastCol)};

        if (lastRow > firstRow)
        {
            CellRangeAddress[] contentRegion = {new CellRangeAddress(firstRow + 1, lastRow, firstCol, lastCol)};
            applyEvenRowStyleRule(contentRegion);
            applyOddRowStyleRule(contentRegion);
        }

        applyHeaderStyleRule(headerRegion);

        return this;
    }

    /**
     * Instantiates a new factory instance initialised with the default {@link ColourStyle#GREY}
     * style.
     *
     * @param tab
     *        the target worksheet
     * @return a new {@code AlternatingRowFactory} instance
     */
    public static AlternatingRowFactory newInstance(Sheet tab)
    {
        return new AlternatingRowFactory(tab, ColourStyle.GREY);
    }

    /**
     * Instantiates a new factory instance initialised with the specified {@link ColourStyle}.
     *
     * @param tab
     *        the target worksheet
     * @param colour
     *        the predefined color scheme to apply
     * @return a new {@code AlternatingRowFactory} instance
     */
    public static AlternatingRowFactory newInstance(Sheet tab, ColourStyle colour)
    {
        return new AlternatingRowFactory(tab, colour);
    }

    /**
     * Validates that RGB integer parameters reside within valid 8-bit bounds (0 to 255).
     *
     * @param red
     *        the red component value
     * @param green
     *        the green component value
     * @param blue
     *        the blue component value
     * @throws IllegalArgumentException
     *         if any component value is less than 0 or greater than 255
     */
    private void validateRGB(int red, int green, int blue)
    {
        if (red < 0 || red > 255 || green < 0 || green > 255 || blue < 0 || blue > 255)
        {
            throw new IllegalArgumentException("RGB components must each be between 0 and 255.");
        }
    }

    /**
     * Resolves a custom RGB colour using a default HSSF palette fallback index.
     *
     * @param red
     *        red component (0 to 255)
     * @param green
     *        green component (0 to 255)
     * @param blue
     *        blue component (0 to 255)
     * @return a valid {@link Color} for the active worksheet type
     */
    private Color getCustomBackgroundRGB(int red, int green, int blue)
    {
        return getCustomBackgroundRGB(red, green, blue, HSSFColor.HSSFColorPredefined.GREY_25_PERCENT);
    }

    /**
     * Resolves a custom RGB colour for the active worksheet type.
     *
     * <p>
     * For HSSF worksheets, the requested colour is looked up in the workbook's custom palette. If
     * it is not already present, the specified palette slot is assigned the requested RGB values
     * and then returned. For XSSF worksheets, an {@link XSSFColor} is created directly from the
     * specified RGB values.
     * </p>
     *
     * @param red
     *        red component, from 0 to 255
     * @param green
     *        green component, from 0 to 255
     * @param blue
     *        blue component, from 0 to 255
     * @param hssfColorIndex
     *        HSSF palette slot to use when the requested colour is not already present in the
     *        custom palette
     * @return a {@link Color} representing the requested RGB colour
     */
    private Color getCustomBackgroundRGB(int red, int green, int blue, HSSFColor.HSSFColorPredefined hssfColorIndex)
    {
        if (sheet instanceof HSSFSheet)
        {
            HSSFPalette palette = ((HSSFWorkbook) sheet.getWorkbook()).getCustomPalette();
            Color ssColor = palette.findColor((byte) red, (byte) green, (byte) blue);

            if (ssColor == null)
            {
                palette.setColorAtIndex(hssfColorIndex.getIndex(), (byte) red, (byte) green, (byte) blue);
                ssColor = palette.getColor(hssfColorIndex.getIndex());
            }

            return ssColor;
        }

        return new XSSFColor(new java.awt.Color(red, green, blue), null);
    }

    /**
     * Applies conditional formatting rules for the header row.
     *
     * @param region
     *        array of {@link CellRangeAddress} defining the header row region
     */
    private void applyHeaderStyleRule(CellRangeAddress[] region)
    {
        ConditionalFormattingRule rule = formatter.createConditionalFormattingRule("ROW()>=0");

        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(headerRowColour);
        fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        FontFormatting font = rule.createFontFormatting();
        font.setFontColorIndex(IndexedColors.WHITE.getIndex());
        font.setFontStyle(false, true); // Bold formatting
        font.setFontHeight((short) (12 * 20));

        setBorderStyleRule(rule);
        formatter.addConditionalFormatting(region, rule);
    }

    /**
     * Applies conditional formatting rules for even-numbered content rows.
     *
     * @param region
     *        array of {@link CellRangeAddress} defining the body content region
     */
    private void applyEvenRowStyleRule(CellRangeAddress[] region)
    {
        ConditionalFormattingRule rule = formatter.createConditionalFormattingRule(EVEN_ROW_RULE);

        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(evenRowColour);
        fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        setBorderStyleRule(rule);
        formatter.addConditionalFormatting(region, rule);
    }

    /**
     * Applies conditional formatting rules for odd-numbered content rows.
     *
     * @param region
     *        array of {@link CellRangeAddress} defining the body content region
     */
    private void applyOddRowStyleRule(CellRangeAddress[] region)
    {
        ConditionalFormattingRule rule = formatter.createConditionalFormattingRule(ODD_ROW_RULE);

        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(oddRowColour);
        fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        setBorderStyleRule(rule);
        formatter.addConditionalFormatting(region, rule);
    }

    /**
     * Configures cell border styles for a given conditional formatting rule.
     *
     * @param rule
     *        the {@link ConditionalFormattingRule} to attach thin borders and border colors to
     */
    private void setBorderStyleRule(ConditionalFormattingRule rule)
    {
        BorderFormatting bordFmt = rule.createBorderFormatting();

        bordFmt.setBorderTop(BorderStyle.THIN);
        bordFmt.setBorderBottom(BorderStyle.THIN);
        bordFmt.setBorderLeft(BorderStyle.THIN);
        bordFmt.setBorderRight(BorderStyle.THIN);

        bordFmt.setTopBorderColor(borderColour);
        bordFmt.setBottomBorderColor(borderColour);
        bordFmt.setLeftBorderColor(borderColour);
        bordFmt.setRightBorderColor(borderColour);
    }

    /**
     * Configures the formatting colours based on a predefined {@link ColourStyle}.
     *
     * @param colour
     *        the predefined color style enum value
     */
    private void setPredefinedColour(ColourStyle colour)
    {
        switch (colour)
        {
            case LIME:
                applyLimeStyles();
            break;

            case BLUE:
                applyBlueStyles();
            break;

            case PURPLE:
                applyPurpleStyles();
            break;

            case CORAL:
                applyCoralStyles();
            break;

            case ORANGE:
                applyOrangeStyles();
            break;

            default:
                applyGreyStyles();
            break;
        }
    }

    /** Applies the predefined Grey colour palette. */
    private void applyGreyStyles()
    {
        /* Windspray Colorbond */
        setHeaderRowRGB(136, 139, 138);

        /* Grey Pail */
        setOddRowRGB(186, 191, 194);

        /* Light Grey */
        setEvenRowRGB(237, 239, 238);

        /* White */
        setBorderRGB(255, 255, 255);
    }

    /** Applies the predefined Sky Blue colour palette. */
    private void applyBlueStyles()
    {
        /* Dark Blue */
        setHeaderRowRGB(79, 129, 189);

        /* White */
        setOddRowRGB(255, 255, 255);

        /* Light Blue */
        setEvenRowRGB(189, 215, 238);

        /* Sky Blue */
        setBorderRGB(149, 179, 215);
    }

    /** Applies the predefined Lime colour palette. */
    private void applyLimeStyles()
    {
        /* Dark Grey */
        setHeaderRowRGB(128, 128, 128);

        /* Lime */
        setOddRowRGB(204, 255, 204);

        /* White */
        setEvenRowRGB(255, 255, 255);

        /* Light Grey */
        setBorderRGB(165, 165, 165);
    }

    /** Applies the predefined Soft Purple colour palette. */
    private void applyPurpleStyles()
    {
        /* Dark Purple */
        setHeaderRowRGB(128, 100, 162);

        /* Middle Purple */
        setOddRowRGB(229, 224, 236);

        /* Pale Purple */
        setEvenRowRGB(204, 192, 218);

        /* White */
        setBorderRGB(255, 255, 255);
    }

    /** Applies the predefined Coral colour palette. */
    private void applyCoralStyles()
    {
        /* Dark Coral */
        setHeaderRowRGB(192, 80, 77);

        /* Pale Coral */
        setOddRowRGB(242, 221, 220);

        /* Middle Coral */
        setEvenRowRGB(230, 185, 184);

        /* White */
        setBorderRGB(255, 255, 255);
    }

    /** Applies the predefined Orange colour palette. */
    private void applyOrangeStyles()
    {
        /* Bright Orange */
        setHeaderRowRGB(237, 125, 49);

        /* Light Grey */
        setOddRowRGB(217, 217, 217);

        /* White */
        setEvenRowRGB(255, 255, 255);

        /* White */
        setBorderRGB(166, 166, 166);
    }
}