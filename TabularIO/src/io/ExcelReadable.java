package io;

/**
 * Extends {@link Readable} with worksheet access and workbook-specific operations for spreadsheet
 * formats that support multiple worksheets.
 *
 * <p>
 * Implementations provide access to worksheet names and counts, and allow individual worksheets to
 * be selected and loaded into memory by name or zero-based index.
 * </p>
 *
 * <b>Change log:</b>
 * <ul>
 * <li>7 October 2026 - Initial creation.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.1
 * @since 7 October 2026
 */
public interface ExcelReadable extends Readable
{
    /**
     * Returns the name of the currently active worksheet.
     *
     * @return name of the active worksheet
     */
    String getWorksheetName();

    /**
     * Returns the total number of worksheets in the workbook.
     *
     * @return total number of worksheets
     */
    int getSheetCount();

    /**
     * Reads and parses the specified worksheet into memory, making it the active worksheet.
     *
     * @param tabName
     *        name of the worksheet to read
     * 
     * @throws TabularDataException
     *         if the worksheet cannot be found, read, or parsed
     */
    void readWorksheet(String tabName) throws TabularDataException;

    /**
     * Reads and parses the worksheet at the specified zero-based index into memory, making it the
     * active worksheet.
     *
     * @param tabIndex
     *        zero-based index of the worksheet to read
     * 
     * @throws TabularDataException
     *         if the worksheet index is invalid or the worksheet cannot be read or parsed
     */
    void readWorksheet(int tabIndex) throws TabularDataException;
}