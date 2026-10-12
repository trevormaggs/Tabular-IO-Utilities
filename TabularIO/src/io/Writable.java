package io;

import java.nio.file.Path;

/**
 * Defines the core contract for writing tabular data to a file.
 *
 * <p>
 * Implementations provide operations for managing row positions, inserting cell values, writing
 * headers and content, and managing the underlying output resource.
 * </p>
 *
 * <p>
 * Implementations may require {@link #open()} to be called before writing data. Resources should be
 * released by calling {@link #close()}.
 * </p>
 *
 * @author Trevor Maggs
 * @version 0.1
 * @since 7 October 2026
 */
public interface Writable extends AutoCloseable
{
    /**
     * Adds a new row to the output data.
     *
     * @return index or position of the newly added row
     */
    int addRow();

    /**
     * Returns the current row position.
     *
     * @return current row index or position
     */
    int getCurrentRowPosition();

    /**
     * Inserts a text value into the current row.
     *
     * @param word
     *        text value to insert
     */
    void insertCell(String word);

    /**
     * Inserts a numeric value into the current row.
     *
     * @param <T>
     *        type of numeric value
     * @param num
     *        numeric value to insert
     */
    <T extends Number> void insertCell(T num);

    /**
     * Writes the specified column headers to the output data.
     *
     * @param headers
     *        array of header values
     * 
     * @throws TabularDataException
     *         if the headers cannot be written
     */
    void writeHeader(String[] headers) throws TabularDataException;

    /**
     * Writes the specified entries as content in the output data.
     *
     * @param entries
     *        array of content values
     * @throws TabularDataException
     * 
     *         if the content cannot be written
     */
    void writeContent(String[] entries) throws TabularDataException;

    /**
     * Returns the path to the output file.
     *
     * @return path to the output file
     */
    Path getFilePath();

    /**
     * Opens and initialises the underlying output resource.
     *
     * @throws TabularDataException
     *         if the resource cannot be opened
     */
    void open() throws TabularDataException;

    /**
     * Closes the underlying output resource and releases associated resources.
     *
     * @throws TabularDataException
     *         if an error occurs while closing the resource
     */
    @Override
    void close() throws TabularDataException;
}