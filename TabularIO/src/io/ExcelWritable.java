package io;

import io.style.AlternatingRowFactory;
import io.style.ColourStyle;

/**
 * Defines an extended writing contract for spreadsheet formats that support multiple worksheets and
 * worksheet formatting.
 *
 * <p>
 * In addition to the operations defined by {@link Writable}, implementations provide functionality
 * for creating worksheets, configuring automatic filters, applying alternating row styles, and
 * accessing the alternating row style
 * factory.
 * </p>
 *
 * @author Trevor Maggs
 * @version 0.1
 * @since 7 October 2026
 */
public interface ExcelWritable extends Writable
{
    /**
     * Creates a new worksheet with the specified tab name.
     *
     * @param tabName
     *        name to assign to the new worksheet
     * 
     * @throws TabularDataException
     *         if the worksheet cannot be created
     */
    void createNewWorksheet(String tabName) throws TabularDataException;

    /**
     * Returns the name of the current worksheet.
     *
     * @return name of the current worksheet
     */
    String getWorksheetName();

    /**
     * Configures an automatic filter for the specified rectangular cell range.
     *
     * @param firstRow
     *        zero-based index of the first row
     * @param firstCol
     *        zero-based index of the first column
     * @param lastRow
     *        zero-based index of the last row
     * @param lastCol
     *        zero-based index of the last column
     */
    void setAutoFilter(int firstRow, int firstCol, int lastRow, int lastCol);

    /**
     * Applies alternating row styles to the specified rectangular cell range.
     *
     * @param firstRow
     *        zero-based index of the first row
     * @param firstCol
     *        zero-based index of the first column
     * @param lastRow
     *        zero-based index of the last row
     * @param lastCol
     *        zero-based index of the last column
     * @param colour
     *        colour style used to configure the alternating row appearance
     */
    void alternateRowStyles(int firstRow, int firstCol, int lastRow, int lastCol, ColourStyle colour);

    /**
     * Returns the alternating row style factory associated with this writer.
     *
     * @return alternating row style factory
     */
    AlternatingRowFactory getInstanceAlternatingRow();
}