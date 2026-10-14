package io.reader;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Parses worksheet XML from an XLSX workbook using SAX events.
 *
 * <p>
 * Cell values are converted to strings and accumulated into row arrays. Shared strings are resolved
 * through the supplied shared-strings table. Numeric values are formatted using worksheet styles
 * when a styles table is available.
 * </p>
 *
 * <p>
 * Formula cells use their cached results stored in the workbook. This handler does not evaluate
 * formulas. Although XML parsing is event-driven, parsed rows are retained in memory in
 * {@link #rowCache}.
 * </p>
 */
class ExcelHandlerSave extends DefaultHandler
{
    private final StylesTable stylesTable;
    private final ReadOnlySharedStringsTable sharedStringsTable;
    private final DataFormatter formatter;
    private final int minColumns;
    private final StringBuilder sbValue;
    private final List<String> cellData;
    private final List<String[]> rowCache;

    private XssfDataType nextDataType;
    private boolean hasCellValue;
    private boolean inlineValue;
    private short formatIndex;
    private String cellRef;
    private String formatString;
    private int maxColumnSize;
    private int currentRow;
    private int currentCol;

    /**
     * Identifies the cell data types handled by this parser.
     */
    private enum XssfDataType
    {
        BOOLEAN, ERROR, FORMULA, INLINE_STRING, SST_STRING, NUMBER
    }

    /**
     * Master constructor initialising all handler dependencies and states.
     *
     * @param sst
     *        shared-strings table used to resolve shared-string cells
     * @param sty
     *        worksheet styles table used to optionally resolve cell formats
     * @param minCol
     *        minimum number of columns in each cached row; non-positive values disable padding
     */
    ExcelHandlerSave(ReadOnlySharedStringsTable sst, StylesTable sty, int minCol)
    {
        this.sharedStringsTable = sst;
        this.stylesTable = sty;
        this.minColumns = Math.max(0, minCol);
        this.cellData = new ArrayList<>();
        this.rowCache = new ArrayList<>();
        this.sbValue = new StringBuilder();
        this.formatter = (sty != null ? new DataFormatter() : null);
        this.maxColumnSize = 0;
        this.currentRow = -1;
        this.currentCol = -1;
    }

    /**
     * Creates a handler without formatting support or minimum row padding.
     *
     * @param sst
     *        shared-strings table used to resolve shared-string cells
     */
    ExcelHandlerSave(ReadOnlySharedStringsTable sst)
    {
        this(sst, null, -1);
    }

    /**
     * Creates a handler with optional minimum row padding.
     *
     * @param sst
     *        shared-strings table used to resolve shared-string cells
     * @param minCol
     *        minimum number of columns in each cached row
     */
    ExcelHandlerSave(ReadOnlySharedStringsTable sst, int minCol)
    {
        this(sst, null, minCol);
    }

    /**
     * Creates a handler with support for cell number and date formatting.
     *
     * @param sst
     *        shared-strings table used to resolve shared-string cells
     * @param sty
     *        worksheet styles table used to resolve cell formats
     */
    ExcelHandlerSave(ReadOnlySharedStringsTable sst, StylesTable sty)
    {
        this(sst, sty, -1);
    }

    /**
     * Determines whether an XML element contains cell text.
     *
     * Important note, 'inlineStr' is only a container, but the actual text is inside 't' when
     * inlineValue is true
     *
     * @param qname
     *        qualified XML element name
     * @return {@code true} if the element contains a cell value or inline-string text
     */

    private boolean textTagVisited(String qname)
    {
        return "v".equals(qname) || ("t".equals(qname) && inlineValue);
    }

    /**
     * Determines whether an XML element represents an inline-string container.
     *
     * @param qname
     *        qualified XML element name
     * @return {@code true} if the element is {@code is}
     */
    private boolean inlineStrTagVisited(String qname)
    {
        return "is".equals(qname);
    }

    /**
     * Determines whether an XML element represents a worksheet row.
     *
     * @param qname
     *        qualified XML element name
     * @return {@code true} if the element is {@code row}
     */
    private boolean rowVisited(String qname)
    {
        return "row".equals(qname);
    }

    /**
     * Determines whether an XML element represents a worksheet cell.
     *
     * @param qname
     *        qualified XML element name
     * @return {@code true} if the element is {@code c}
     */
    private boolean cellVisited(String qname)
    {
        return "c".equals(qname);
    }

    /**
     * Processes the start of an XML element.
     *
     * <p>
     * Initialises cell parsing state, identifies cell data types, and resolves number-format
     * information when a styles table is available.
     * </p>
     *
     * @param uri
     *        namespace URI of the element
     * @param localName
     *        local element name
     * @param qname
     *        qualified element name
     * @param attributes
     *        attributes associated with the element
     *
     * @throws SAXException
     *         if SAX processing fails or a row reference is invalid
     */
    @Override
    public void startElement(String uri, String localName, String qname, Attributes attributes) throws SAXException
    {
        if (textTagVisited(qname))
        {
            hasCellValue = true;
            sbValue.setLength(0);
        }

        else if (inlineStrTagVisited(qname))
        {
            inlineValue = true;
        }

        else if (rowVisited(qname))
        {
            String rowRef = attributes.getValue("r");

            currentCol = -1;
            cellData.clear();

            try
            {
                currentRow = (rowRef != null ? Integer.parseInt(rowRef) - 1 : currentRow + 1);
            }

            catch (NumberFormatException exc)
            {
                throw new SAXException("Invalid worksheet row reference: " + rowRef, exc);
            }
        }

        else if (cellVisited(qname))
        {
            String cellType = attributes.getValue("t");
            String cellStyle = attributes.getValue("s");

            // Hard-reset cell-specific state to prevent state leakage between cells
            hasCellValue = false;
            sbValue.setLength(0);
            formatIndex = -1;
            formatString = null;
            cellRef = attributes.getValue("r");

            if ("b".equals(cellType))
            {
                nextDataType = XssfDataType.BOOLEAN;
            }

            else if ("e".equals(cellType))
            {
                nextDataType = XssfDataType.ERROR;
            }

            else if ("inlineStr".equals(cellType))
            {
                nextDataType = XssfDataType.INLINE_STRING;
            }

            else if ("s".equals(cellType))
            {
                nextDataType = XssfDataType.SST_STRING;
            }

            else if ("str".equals(cellType))
            {
                nextDataType = XssfDataType.FORMULA;
            }

            else
            {
                nextDataType = XssfDataType.NUMBER;

                if (stylesTable != null && cellStyle != null)
                {
                    readCellFormat(cellStyle);
                }
            }
        }
    }

    /**
     * Appends a SAX character-data chunk to the current cell value.
     *
     * <p>
     * SAX may deliver the text of a single XML element in multiple chunks, so each chunk is
     * appended to the buffer.
     * </p>
     *
     * @param ch
     *        character buffer supplied by the SAX parser
     * @param start
     *        offset of the first character
     * @param length
     *        number of characters to append
     * @throws SAXException
     *         if SAX processing fails
     */
    @Override
    public void characters(char[] ch, int start, int length) throws SAXException
    {
        if (hasCellValue)
        {
            sbValue.append(ch, start, length);
        }
    }

    /**
     * Processes the end of an XML element.
     *
     * <p>
     * Cell values are resolved and appended to the current row. At the end of a row, trailing
     * padding is added when required and the completed row is cached.
     * </p>
     *
     * @param uri
     *        namespace URI of the element
     * @param localName
     *        local element name
     * @param qname
     *        qualified element name
     * @throws SAXException
     *         if SAX processing fails or a cell reference is invalid
     */
    @Override
    public void endElement(String uri, String localName, String qname) throws SAXException
    {
        if ("v".equals(qname) && !inlineValue)
        {
            // A normal cell value is complete.
            appendCellValue();
            hasCellValue = false;
        }

        else if (inlineStrTagVisited(qname))
        {
            // All text runs in the inline string have been accumulated.
            appendCellValue();
            inlineValue = false;
            hasCellValue = false;
        }

        else if (rowVisited(qname))
        {
            // minColumns is a count. currentCol is a zero-based index.
            while (cellData.size() < minColumns)
            {
                cellData.add("");
            }

            if (!cellData.isEmpty())
            {
                rowCache.add(cellData.toArray(new String[0]));
            }

            cellData.clear();
            currentCol = -1;
        }
    }

    public void clearCachedRow()
    {
        rowCache.clear();
    }

    public int getCachedRowHeight()
    {
        return rowCache.size();
    }

    /**
     * Returns the maximum number of columns detected across the worksheet. This value can be used
     * to determine a consistent matrix width when processing the parsed rows.
     *
     * @return maximum number of columns detected, or {@code 0} if no cell values were processed
     */
    public int getMaximumMatrixWidth()
    {
        return maxColumnSize;
    }

    /**
     * Returns an iterator over the cached row arrays.
     *
     * @return row iterator
     */
    public Iterator<String[]> rowIterator()
    {
        return rowCache.iterator();
    }

    /**
     * Resolves a cell's style index into its number-format index and format string.
     *
     * <p>
     * If the style cannot be resolved, the raw cell value is retained rather than applying a
     * potentially incorrect format.
     * </p>
     *
     * <p>
     * It also converts common US-style date patterns to a standardised Australian day/month/year
     * time format.
     * </p>
     *
     * @param styleText
     *        style index from the cell's {@code s} attribute
     */
    private void readCellFormat(String styleText)
    {
        try
        {
            int styleIndex = Integer.parseInt(styleText);
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);

            formatIndex = style.getDataFormat();
            formatString = style.getDataFormatString();

            if (formatString == null)
            {
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }

            // Convert common US-style date patterns to day/month order.
            if ("m/d/yy".equalsIgnoreCase(formatString))
            {
                formatString = "d/m/yy";
            }

            else if ("m/d/yy h:mm".equalsIgnoreCase(formatString))
            {
                formatString = "d/m/yy h:mm";
            }
        }

        catch (IllegalArgumentException | IndexOutOfBoundsException exc)
        {
            formatString = null;
        }
    }

    /**
     * Resolves the current cell value and appends it to the current row.
     *
     * @throws SAXException
     *         if the cell reference is invalid or the cells
     *         are not in ascending column order
     */
    private void appendCellValue() throws SAXException
    {
        String value = resolveCellValue();
        int thisCol = (cellRef != null ? nameToColumn(cellRef) : currentCol + 1);

        if (thisCol < 0)
        {
            throw new SAXException("Invalid worksheet cell reference: " + cellRef);
        }

        if (thisCol <= currentCol)
        {
            throw new SAXException("Worksheet cells are not in ascending column order: " + cellRef);
        }

        // Preserve the positions of cells omitted from the worksheet XML.
        int missingCells = thisCol - currentCol - 1;

        for (int i = 0; i < missingCells; i++)
        {
            cellData.add("");
        }

        cellData.add(value);
        currentCol = thisCol;
        maxColumnSize = Math.max(maxColumnSize, thisCol + 1);
    }

    /**
     * Converts the current cell's XML text into a string value.
     *
     * @return resolved cell value, or an empty string if the value is {@code null} or empty
     *
     * @throws SAXException
     *         if a shared-string index is invalid or the shared-strings table is unavailable
     */
    private String resolveCellValue() throws SAXException
    {
        String value;
        String rawValue = sbValue.toString();

        switch (nextDataType)
        {
            case BOOLEAN:
                value = (!rawValue.isEmpty() && rawValue.charAt(0) == '0' ? "FALSE" : "TRUE");
            break;

            case ERROR:
                value = "ERROR:" + rawValue;
            break;

            case FORMULA:
            case INLINE_STRING:
                value = rawValue;
            break;

            case SST_STRING:
                try
                {
                    int index = Integer.parseInt(rawValue);

                    if (sharedStringsTable == null)
                    {
                        throw new SAXException("Shared-strings table is unavailable for index " + index);
                    }

                    value = sharedStringsTable.getItemAt(index).getString();
                }

                catch (NumberFormatException | IndexOutOfBoundsException exc)
                {
                    // Revert to the raw value as a fallback when shared string tables are not
                    // matched
                    value = rawValue;
                }
            break;

            case NUMBER:
                value = formatNumber(rawValue);
            break;

            default:
                value = rawValue;
            break;
        }

        return (value == null ? "" : value.trim());
    }

    /**
     * Formats a numeric cell value using its worksheet style, when available.
     *
     * @param rawValue
     *        numeric text from the worksheet XML
     * @return formatted value, or the original text if formatting is unavailable or the value
     *         cannot be parsed
     */
    private String formatNumber(String rawValue)
    {
        if (stylesTable == null || formatter == null || formatString == null || rawValue.isEmpty())
        {
            return rawValue;
        }

        try
        {
            return formatter.formatRawCellContents(Double.parseDouble(rawValue), formatIndex, formatString);
        }

        catch (NumberFormatException exc)
        {
            return rawValue;
        }
    }

    /**
     * Converts an Excel cell reference, such as {@code A1} or {@code AB12}, into a zero-based
     * column index.
     *
     * @param ref
     *        cell reference containing column letters and a row number
     * @return zero-based column index, or {@code -1} if no valid column index can be determined
     */
    private static int nameToColumn(String ref)
    {
        if (ref != null && !ref.isEmpty())
        {
            int col = 0;
            int letterCount = 0;

            for (int i = 0; i < ref.length(); i++)
            {
                char c = ref.charAt(i);

                if (Character.isDigit(c))
                {
                    break;
                }

                if (c >= 'A' && c <= 'Z')
                {
                    col = col * 26 + (c - 'A' + 1);
                    letterCount++;
                }

                else if (c >= 'a' && c <= 'z')
                {
                    col = col * 26 + (c - 'a' + 1);
                    letterCount++;
                }

                else
                {
                    return -1;
                }
            }

            return (letterCount == 0 ? -1 : col - 1);
        }

        return -1;
    }
}