package io;

import java.nio.file.Path;

/**
 * Defines the contract for reading tabular data sources, such as spreadsheets and CSV files.
 *
 * <p>
 * Implementations provide mechanisms to open and close data resources, inspect data dimensions,
 * locate columns by cell content or regular expression, and iterate over rows represented as arrays
 * of strings.
 * </p>
 *
 * <p>
 * Implementations may require {@link #open()} to be called before data can be accessed. Resources
 * should be released by calling {@link #close()}.
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
public interface Readable extends AutoCloseable, Iterable<String[]>
{
    /**
     * Returns the path to the target data file.
     *
     * @return path to the target data file
     */
    Path getFilePath();

    /**
     * Opens and initialises the underlying data resource for reading.
     *
     * @throws TabularDataException
     *         if the resource cannot be opened or parsed
     */
    void open() throws TabularDataException;

    /**
     * Closes the underlying resource and releases associated resources, such as open file handles
     * or buffers.
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
}