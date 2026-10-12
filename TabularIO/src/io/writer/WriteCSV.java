package io.writer;

import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import io.TabularDataException;
import io.Writable;
import io.TabularDataException.ErrorCode;
import util.Separator;

/**
 * Implements the {@link Writable} interface to create CSV files using field delimiters and quoted
 * fields.
 * 
 * <ul>
 * <li>Released by Trevor Maggs on 15 April 2016</li>
 * <li>Modified to support for JDK1.7 NIO.2 functionality - 1 May 2017</li>
 * <li>Major improvements for efficient memory management - 22 November 2017</li>
 * <li>Fixed bug in addRow method - 23 January 2018</li>
 * <li>Flush method added to improve efficiency - 10 August 2020</li>
 * </ul>
 * 
 * @author Trevor Maggs
 * @version 0.6
 */
public class WriteCSV implements Writable
{
    private static final char DEFAULT_DELIMITER = ',';
    private static final char EXCEL_ENCLOSURE = '"';
    private final List<List<String>> csvData;
    private int colCounter;
    private int maxColCount;
    protected final Path filePath;
    protected BufferedWriter csvWriter;

    /**
     * Default constructor to create a CSV file with the default filename.
     */
    public WriteCSV()
    {
        this("out.csv");
    }

    /**
     * Constructs a CSV writer targeting the specified file path.
     * 
     * @param file
     *        relative or absolute path to the file
     */
    public WriteCSV(String file)
    {
        if (file != null && !file.isEmpty())
        {
            Path fpath = Paths.get(file);
            Path fileNamePath = fpath.getFileName();

            if (fileNamePath != null)
            {
                String filename = fileNamePath.toString();
                int extPos = filename.lastIndexOf('.');

                if (extPos >= 0)
                {
                    String normalised = filename.substring(0, extPos) + ".csv";
                    Path parent = fpath.getParent();

                    this.filePath = (parent != null ? parent.resolve(normalised) : Paths.get(normalised));
                }

                else
                {
                    this.filePath = fpath.resolveSibling(filename + ".csv");
                }
            }

            else
            {
                // Fallback for root paths without filenames, for example: "/" or "C:\"
                String filename = fpath.toString();
                int extPos = filename.lastIndexOf('.');

                if (extPos >= 0)
                {
                    this.filePath = Paths.get(filename.substring(0, extPos) + ".csv");
                }

                else
                {
                    this.filePath = Paths.get(filename + ".csv");
                }
            }
        }

        else
        {
            this.filePath = Paths.get("out.csv");
        }

        this.csvData = new ArrayList<>();
    }

    /**
     * Adds a new empty row to the CSV output.
     *
     * @return the zero-based index of the newly added row
     */
    @Override
    public int addRow()
    {
        colCounter = 0;
        csvData.add(new ArrayList<>());

        return getCurrentRowPosition();
    }

    /**
     * Returns the zero-based index of the current row.
     *
     * @return the zero-based index of the current row, or {@code -1} if no rows exist
     */
    @Override
    public int getCurrentRowPosition()
    {
        return csvData.size() - 1;
    }

    /**
     * Adds a string value to the current row.
     *
     * <p>
     * If no row currently exists, a new row is created automatically.
     * </p>
     *
     * @param word
     *        string value to add to the current row
     */
    @Override
    public void insertCell(String word)
    {
        if (csvData.isEmpty())
        {
            addRow();
        }

        List<String> csvLine = csvData.get(csvData.size() - 1);

        csvLine.add(word);
        colCounter++;

        if (colCounter > maxColCount)
        {
            maxColCount = colCounter;
        }
    }

    /**
     * Adds a numeric value to the current row.
     *
     * <p>
     * A {@code null} value is written as an empty field.
     * </p>
     *
     * @param num
     *        numeric value to add to the current row
     * @param <T>
     *        numeric type
     */
    @Override
    public <T extends Number> void insertCell(T num)
    {
        insertCell(num != null ? num.toString() : "");
    }

    /**
     * Adds the supplied values as the first row of the CSV output.
     *
     * @param headers
     *        column headings to add as a row
     */
    @Override
    public void writeHeader(String[] headers)
    {
        writeContent(headers);
    }

    /**
     * Adds an array of values as a complete row.
     *
     * <p>
     * A {@code null} array is ignored. The maximum column count is updated when the supplied row
     * contains more columns than any previously added row.
     * </p>
     *
     * @param entries
     *        values to add as a row
     */
    @Override
    public void writeContent(String[] entries)
    {
        if (entries != null)
        {
            if (entries.length > maxColCount)
            {
                maxColCount = entries.length;
            }

            colCounter = entries.length;
            csvData.add(new ArrayList<>(Arrays.asList(entries)));
        }
    }

    @Override
    public Path getFilePath()
    {
        return filePath;
    }

    @Override
    public void open() throws TabularDataException
    {
        if (!filePath.toString().endsWith(".csv"))
        {
            throw new TabularDataException("The file [" + filePath + "] does not appear to be a valid CSV file.");
        }

        try
        {
            csvWriter = Files.newBufferedWriter(filePath, StandardCharsets.UTF_8);
        }

        catch (IOException exc)
        {
            throw new TabularDataException(ErrorCode.FILE_ACCESS_DENIED, "Cannot open CSV file [" + filePath + "]. Check permissions.", exc);
        }
    }

    @Override
    public void close() throws TabularDataException
    {
        try
        {
            writeBufferedData();
        }

        finally
        {
            if (csvWriter != null)
            {
                try
                {
                    csvWriter.flush();
                    csvWriter.close();
                }

                catch (IOException exc)
                {
                    throw new TabularDataException(ErrorCode.STREAM_WRITE_FAILED, "Problem while attempting to close CSV file [" + filePath + "]", exc);
                }
            }
        }
    }

    /**
     * Writes accumulated buffered rows to disk. Called internally prior to closing resources.
     * 
     * @throws TabularDataException
     *         if an I/O error occurs
     */
    private void writeBufferedData() throws TabularDataException
    {
        if (csvWriter == null)
        {
            return;
        }

        StringBuilder buffer = new StringBuilder();
        Separator sep = new Separator(String.valueOf(DEFAULT_DELIMITER));

        try
        {
            for (List<String> row : csvData)
            {
                for (String str : row)
                {
                    buffer.append(sep).append(escapeEmbeddedCharacters(str));
                }

                // Pad the row with additional blank cells if the row is shorter than the maximum
                // number of columns. That is necessary to strictly maintain a matrix spreadsheet.
                int missingCols = maxColCount - row.size();

                for (int i = 0; i < missingCols; i++)
                {
                    buffer.append(sep);
                }

                csvWriter.write(buffer.toString());
                csvWriter.newLine();

                buffer.setLength(0);
                sep.reset();
            }
        }

        catch (IOException exc)
        {
            throw new TabularDataException(ErrorCode.STREAM_WRITE_FAILED, "There was a problem while writing the data to file [" + filePath + "]", exc);
        }
    }

    /**
     * Normalises, escapes, and quotes a CSV field when its contents require special handling,
     * stripping leading and trailing whitespace by default.
     *
     * @param value
     *        field value to normalise and escape
     * @return the CSV-safe representation of the field, or an empty string if {@code value} is
     *         {@code null}
     */
    private String escapeEmbeddedCharacters(String value)
    {
        return escapeEmbeddedCharacters(value, true);
    }

    /**
     * Normalises, escapes, and quotes a CSV field when its contents require special handling.
     *
     * <p>
     * Leading and trailing whitespace is conditionally removed before the field is evaluated for
     * CSV escaping. Fields containing quotation marks, the delimiter, or line-break characters
     * are enclosed in double quotes. Embedded quotation marks are represented by two consecutive
     * quotation marks.
     * </p>
     *
     * @param value
     *        field value to normalise and escape
     * @param trimmed
     *        {@code true} to strip leading and trailing whitespace; {@code false} to preserve it
     * @return the CSV-safe representation of the field, or an empty string if {@code value} is
     *         {@code null}
     */
    private String escapeEmbeddedCharacters(String value, boolean trimmed)
    {
        if (value != null)
        {
            value = (trimmed ? value.trim() : value);

            boolean hasQuotes = value.indexOf('"') >= 0;
            boolean hasDelimiter = value.indexOf(DEFAULT_DELIMITER) >= 0
                    || value.indexOf('\n') >= 0
                    || value.indexOf('\r') >= 0;

            if (!hasQuotes && !hasDelimiter)
            {
                return value;
            }

            StringBuilder buffer = new StringBuilder().append(EXCEL_ENCLOSURE);
            buffer.append(value.replace("\"", "\"\""));
            buffer.append(EXCEL_ENCLOSURE);

            return buffer.toString();
        }

        return "";
    }
}