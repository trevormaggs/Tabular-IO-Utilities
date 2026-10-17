package io;

import java.nio.file.Path;
import java.util.List;

/**
 * Defines the contract for reading tabular data sources, such as spreadsheets and CSV files.
 *
 * <p>
 * Implementations provide mechanisms to inspect data dimensions, locate columns by cell content or
 * regular expression, extract column values, and iterate over rows represented as arrays of
 * strings.
 * </p>
 *
 * <p>
 * Data resources are opened automatically upon construction and should be released by closing the
 * resource. A {@code try-with-resources} statement is strongly recommended.
 * </p>
 *
 * <b>Change log:</b>
 * <ul>
 * <li>7 October 2026 - Initial creation.</li>
 * <li>17 October 2026 - Removed explicit open() method to support auto-opening constructors and
 * added column values extraction contracts.</li>
 * </ul>
 * 
 * @author Trevor Maggs
 * @version 0.2
 * @since 7 October 2026
 */
public interface Readable extends AutoCloseable, Iterable<String[]>
{
    /**
     * Returns the path to the target data file.
     *
     * @return path to the target data file
     */
    Path getFilePath();

    /**
     * Closes the reader and releases any associated resources. Calling this method more than once
     * has no additional effect.
     *
     * @throws TabularDataException
     *         if an error occurs while closing the resource
     */
    @Override
    void close() throws TabularDataException;

    /**
     * Returns the number of rows in the data source.
     *
     * @return number of rows
     */
    int getRowCount();

    /**
     * Returns the number of columns in the data source.
     *
     * @return number of columns
     */
    int getColumnCount();

    /**
     * Finds the first column whose cell content matches the specified entry in the given row.
     *
     * @param entry
     *        cell content or header name to locate
     * @param rowIndex
     *        zero-based row index to search
     * @return zero-based column index, or {@code -1} if no match is found
     */
    int findColumnIndexByContent(String entry, int rowIndex);

    /**
     * Finds the first column whose cell content matches the specified entry in the first row (index
     * {@code 0}).
     *
     * @param entry
     *        cell content or header name to locate
     * @return zero-based column index, or {@code -1} if no match is found
     */
    default int findColumnIndexByContent(String entry)
    {
        return findColumnIndexByContent(entry, 0);
    }

    /**
     * Finds the first column whose cell content matches the specified regular expression in the
     * given row.
     *
     * @param regex
     *        regular expression used to match cell content
     * @param rowIndex
     *        zero-based row index to search
     * @return zero-based column index, or {@code -1} if no match is found
     */
    int findColumnIndexRegex(String regex, int rowIndex);

    /**
     * Finds the first column whose cell content matches the specified regular expression in the
     * first row (index {@code 0}).
     *
     * @param regex
     *        regular expression used to match cell content
     * @return zero-based column index, or {@code -1} if no match is found
     */
    default int findColumnIndexRegex(String regex)
    {
        return findColumnIndexRegex(regex, 0);
    }

    /**
     * Returns all cell values for the specified column index across the data source.
     *
     * @param columnIndex
     *        zero-based column index
     * @return list of cell values as strings for the specified column
     */
    List<String> getColumnValues(int columnIndex);

    /**
     * Returns all cell values for the column matching the specified header name.
     *
     * @param headerName
     *        name of the column header to locate (case-insensitive)
     * @return list of cell values for that column, or an empty list if not found
     */
    List<String> getColumnValues(String headerName);
}