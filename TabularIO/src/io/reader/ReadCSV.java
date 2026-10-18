package io.reader;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.function.BiConsumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;
import io.Readable;
import io.TabularDataException;
import io.TabularDataException.ErrorCode;

/**
 * Reads CSV files into memory using a quote-aware parser that supports quoted fields, escaped
 * quotes, and multiline records.
 *
 * <p>
 * The CSV file is parsed automatically when an instance is constructed. Parsed records are retained
 * in memory until {@link #close()} is called. The class provides tabular access through
 * {@link Readable} and supports row iteration through its iterator implementation.
 * </p>
 *
 * <p>
 * A {@code try-with-resources} statement is strongly recommended when processing is complete,
 * unless {@link #close()} is called explicitly.
 * </p>
 * 
 * <b>Change log</b>
 *
 * <ul>
 * <li>Created by Trevor Maggs on 13 November 2024</li>
 * <li>Overhauled by Trevor Maggs on 18 October 2025 to assert Josh Bloch's Effective Java book,
 * "Design and document for inheritance or else prohibit it" (Item 19), therefore it favours
 * composition over inheritance to prevent complexity (Item 18)</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.7
 * @since 7 February 2025
 */
public final class ReadCSV implements Readable
{
    private final Path filePath;
    private final List<String[]> dataset;
    private final BiConsumer<String[], Integer> rowProcessor;
    private int lineNumber;
    private int maxColumnSize;

    /**
     * Constructs and parses a {@code ReadCSV} instance targeting the specified file path, using a
     * custom row processor callback for streaming row/header inspection if needed.
     *
     * @param csvFile
     *        path to the target CSV file
     * @param rowProcessor
     *        an optional {@link BiConsumer} callback accepting row tokens and line numbers during
     *        parsing
     *
     * @throws TabularDataException
     *         if the file path is null, does not exist, is not a regular file, or parsing fails
     */
    public ReadCSV(Path csvFile, BiConsumer<String[], Integer> rowProcessor) throws TabularDataException
    {
        if (csvFile == null)
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "Target path cannot be null.");
        }

        if (!Files.isRegularFile(csvFile))
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "File [" + csvFile + "] is not a regular file or does not exist.");
        }

        String fname = csvFile.toString().toLowerCase();

        if (!fname.endsWith(".csv"))
        {
            throw new TabularDataException(ErrorCode.UNSUPPORTED_FORMAT, "File [" + csvFile + "] does not appear to be a valid CSV matrix format.");
        }

        this.maxColumnSize = 0;
        this.filePath = csvFile;
        this.dataset = new ArrayList<>();
        this.rowProcessor = rowProcessor != null ? rowProcessor : new BiConsumer<String[], Integer>()
        {
            @Override
            public void accept(String[] rowTokens, Integer lineNumber)
            {
                // Do nothing unless specified
            }
        };

        parse();
    }

    /**
     * Constructs and parses a {@code ReadCSV} instance targeting the specified file path.
     *
     * @param csvFile
     *        path to the target CSV file
     *
     * @throws InvalidPathException
     *         if the specified path string is invalid
     * @throws TabularDataException
     *         if the path is null, does not identify a regular file,
     *         has an unsupported file extension, or the CSV cannot be parsed
     */
    public ReadCSV(Path csvFile) throws TabularDataException
    {
        this(csvFile, null);
    }

    /**
     * Constructs and parses a {@code ReadCSV} instance for the specified CSV file path string
     * with default no-op row processing.
     *
     * @param csvFile
     *        path to the target CSV file in string form
     *
     * @throws InvalidPathException
     *         if the specified path string is invalid
     * @throws TabularDataException
     *         if the path is null, does not identify a regular file,
     *         has an unsupported file extension, or the CSV cannot be parsed
     */
    public ReadCSV(String csvFile) throws TabularDataException
    {
        this(csvFile == null ? null : Paths.get(csvFile));
    }

    /**
     * Constructs and parses a {@code ReadCSV} instance for the specified CSV file path string
     * and custom row processor callback.
     *
     * @param csvFile
     *        path to the target CSV file in string form
     * @param rowProcessor
     *        a {@link BiConsumer} callback accepting row tokens and line numbers during parsing
     *
     * @throws InvalidPathException
     *         if the specified path string is invalid
     * @throws TabularDataException
     *         if the path is null, does not identify a regular file,
     *         has an unsupported file extension, or the CSV cannot be parsed
     */
    public ReadCSV(String csvFile, BiConsumer<String[], Integer> rowProcessor) throws TabularDataException
    {
        this(csvFile == null ? null : Paths.get(csvFile), rowProcessor);
    }

    /**
     * Returns an iterator over row token arrays.
     *
     * @return row iterator
     */
    @Override
    public Iterator<String[]> iterator()
    {
        return new Iterator<String[]>()
        {
            private int index = 0;

            @Override
            public boolean hasNext()
            {
                return index < getRowCount();
            }

            @Override
            public String[] next()
            {
                if (!hasNext())
                {
                    throw new NoSuchElementException();
                }

                return dataset.get(index++);
            }
        };
    }

    /**
     * Returns the file path of the CSV document.
     *
     * @return path to the target file
     */
    @Override
    public Path getFilePath()
    {
        return filePath;
    }

    /**
     * Releases all parsed row data held in memory.
     */
    @Override
    public void close()
    {
        dataset.clear();
        maxColumnSize = 0;
    }

    /**
     * Returns the total number of loaded rows (including the header row).
     *
     * @return total row count
     */
    @Override
    public int getRowCount()
    {
        return dataset.size();
    }

    /**
     * Returns the maximum column count calculated across the entire document matrix.
     *
     * @return maximum column count across all parsed rows
     */
    @Override
    public int getColumnCount()
    {
        return maxColumnSize;
    }

    /**
     * Finds the column index matching exact content in the specified row position.
     *
     * @param entry
     *        cell text to find
     * @param rowIndex
     *        zero-based row position
     *
     * @return zero-based column index, or -1 if not found
     */
    @Override
    public int findColumnIndexByContent(String entry, int rowIndex)
    {
        if (entry != null)
        {
            return findColumnIndexRegex(String.format("^%s$", Pattern.quote(entry.trim())), rowIndex);
        }

        return -1;
    }

    /**
     * Finds the column index matching a regular expression in the specified row position.
     *
     * @param regex
     *        regular expression pattern
     * @param rowIndex
     *        zero-based row position
     *
     * @return zero-based column index, or -1 if not found
     */
    @Override
    public int findColumnIndexRegex(String regex, int rowIndex)
    {
        if (regex == null || regex.isEmpty() || rowIndex < 0 || rowIndex >= getRowCount())
        {
            return -1;
        }

        Pattern pattern;
        String[] row = dataset.get(rowIndex);

        try
        {
            pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
        }

        catch (PatternSyntaxException exc)
        {
            pattern = Pattern.compile(Pattern.quote(regex), Pattern.CASE_INSENSITIVE);
        }

        for (int j = 0; j < row.length; j++)
        {
            Matcher matcher = pattern.matcher(row[j] != null ? row[j].trim() : "");

            if (matcher.matches())
            {
                return j;
            }
        }

        return -1;
    }

    /**
     * Returns all cell values for the specified column index across the CSV dataset.
     *
     * @param columnIndex
     *        zero-based column index
     * @return list of cell values as strings for the specified column
     */
    @Override
    public List<String> getColumnValues(int columnIndex)
    {
        List<String> values = new ArrayList<>();

        if (columnIndex >= 0 && columnIndex < getColumnCount())
        {
            int rowCount = getRowCount();

            for (int i = 0; i < rowCount; i++)
            {
                String[] row = dataset.get(i);

                if (row != null && columnIndex < row.length)
                {
                    values.add(row[columnIndex] != null ? row[columnIndex] : "");
                }
                
                else
                {
                    values.add("");
                }
            }
        }

        return values;
    }

    /**
     * Returns the data values for the column identified by its header.
     *
     * <p>
     * The first record is treated as the header record. Matching is case-insensitive, and the
     * header itself is excluded from the result.
     * </p>
     *
     * @param headerName
     *        column header to locate
     * @return data values for the matching column, or an empty list if no matching header exists
     * 
     * @throws NullPointerException
     *         if headerName is null
     */
    @Override
    public List<String> getColumnValues(String headerName)
    {
        headerName = Objects.requireNonNull(headerName, "Header cannot be null").trim();

        if (!headerName.isEmpty())
        {
            int colIndex = findColumnIndexByContent(headerName, 0);

            if (colIndex != -1)
            {
                List<String> values = getColumnValues(colIndex);

                if (!values.isEmpty())
                {
                    // Removes the header text itself, leaving only data values
                    values.remove(0);
                }

                return values;
            }
        }

        return new ArrayList<>();
    }

    /**
     * Returns a string cell value at the specified row and column coordinates.
     *
     * @param row
     *        zero-based row index
     * @param col
     *        zero-based column index
     *
     * @return cell value string, or empty string if out of bounds
     */
    public String getCellValue(int row, int col)
    {
        if (row >= 0 && row < getRowCount())
        {
            String[] r = dataset.get(row);

            if (col >= 0 && col < r.length)
            {
                return r[col] != null ? r[col] : "";
            }
        }

        return "";
    }

    /**
     * Parses the CSV file contents into memory, tokenising multi-line records and facilitating
     * sanitisation and callback processing.
     *
     * @throws TabularDataException
     *         if the file cannot be accessed, does not exist, or an I/O error occurs
     */
    private void parse() throws TabularDataException
    {
        dataset.clear();
        maxColumnSize = 0;

        try (BufferedReader br = Files.newBufferedReader(filePath, StandardCharsets.UTF_8))
        {
            lineNumber = 0;
            String[] record;

            while ((record = parseDataCSV(br)) != null)
            {
                sanitise(record);
                rowProcessor.accept(record, lineNumber);

                // Find the maximum width for uniform matrix across all rows
                maxColumnSize = Math.max(maxColumnSize, record.length);

                lineNumber++;
                dataset.add(record);
            }

            normaliseMatrixWidth();
        }

        catch (NoSuchFileException exc)
        {
            dataset.clear();
            maxColumnSize = 0;
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "File [" + filePath + "] does not exist.", exc);
        }

        catch (IOException exc)
        {
            dataset.clear();
            maxColumnSize = 0;
            throw new TabularDataException(ErrorCode.STREAM_READ_FAILED, "An error occurred while reading CSV file [" + filePath + "].", exc);
        }
    }

    /**
     * Sanitises the values in a parsed CSV record.
     *
     * <p>
     * Replaces each consecutive sequence of carriage return and line feed characters within a cell
     * value with a single space. Other whitespace, including leading and trailing spaces, is
     * preserved.
     * </p>
     *
     * @param rowTokens
     *        array of parsed cell values for the current record
     */
    private void sanitise(String[] rowTokens)
    {
        for (int i = 0; i < rowTokens.length; i++)
        {
            if (rowTokens[i] != null)
            {
                rowTokens[i] = rowTokens[i].replaceAll("[\\r\\n]+", " ");
            }
        }
    }

    /**
     * Parses a single CSV record from a {@link BufferedReader}, supporting b quoted fields, escaped
     * double quotes, and embedded line breaks.
     *
     * <p>
     * Leading whitespace outside quoted fields is skipped, and trailing whitespace is removed from
     * each field. Whitespace handling depends on {@code isWhitespace(char)}. This parser is
     * permissive and does not enforce all RFC 4180 syntax rules.
     * </p>
     *
     * @param br
     *        reader positioned at the start of the next CSV record
     * @return parsed field values, or {@code null} if EOF is reached before any record is read
     *
     * @throws IOException
     *         if an unterminated quoted field is encountered at EOF or an I/O error occurs
     */
    private String[] parseDataCSV(BufferedReader br) throws IOException
    {
        String line;
        boolean inQuotes = false;
        boolean fieldStart = true;
        StringBuilder sb = new StringBuilder();
        List<String> tokens = new ArrayList<>();

        while ((line = br.readLine()) != null)
        {
            for (int i = 0; i < line.length(); i++)
            {
                char c = line.charAt(i);

                // Skip leading unquoted whitespace
                if (fieldStart && !inQuotes && isWhitespace(c))
                {
                    continue;
                }

                if (c == '"')
                {
                    if (!inQuotes)
                    {
                        if (sb.length() == 0)
                        {
                            inQuotes = true;
                        }

                        else
                        {
                            sb.append(c);
                        }
                    }

                    else
                    {
                        // Check for doubled-up escaped quote ("")
                        if (i + 1 < line.length() && line.charAt(i + 1) == '"')
                        {
                            sb.append('"');
                            i++;
                        }

                        else
                        {
                            inQuotes = false;
                        }
                    }

                    fieldStart = false;
                }

                else if (c == ',' && !inQuotes)
                {
                    // Clean up trailing unquoted spaces
                    tokens.add(trimTrailing(sb.toString()));
                    sb.setLength(0);
                    fieldStart = true;
                }

                else
                {
                    sb.append(c);
                    fieldStart = false;
                }
            }

            if (!inQuotes)
            {
                tokens.add(trimTrailing(sb.toString()));
                return tokens.toArray(new String[0]);
            }

            sb.append('\n');
        }

        // Reached end-of-file inside quotes
        if (inQuotes)
        {
            throw new IOException("Malformed CSV: Unterminated double-quoted field encountered.");
        }

        // Stream ended
        return null;
    }

    /**
     * Pads all rows in the dataset to match the maximum column width across the file.
     */
    private void normaliseMatrixWidth()
    {
        if (maxColumnSize > 0 && !dataset.isEmpty())
        {
            for (int i = 0; i < dataset.size(); i++)
            {
                String[] row = dataset.get(i);

                if (row.length < maxColumnSize)
                {
                    String[] replaceRow = new String[maxColumnSize];
                    System.arraycopy(row, 0, replaceRow, 0, row.length);

                    for (int j = row.length; j < maxColumnSize; j++)
                    {
                        replaceRow[j] = "";
                    }

                    dataset.set(i, replaceRow);
                }
            }
        }
    }

    /**
     * Evaluates whether a character represents ASCII whitespace, Java Unicode whitespace,
     * or a Non-Breaking Space ({@code \u00A0}).
     *
     * @param c
     *        character to test
     * @return {@code true} if the character is considered whitespace, otherwise {@code false}
     */
    private static boolean isWhitespace(char c)
    {
        return c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\u00A0' || Character.isWhitespace(c);
    }

    /**
     * Fast, allocation-free trailing trim for field tokens.
     *
     * @param input
     *        string token to trim
     * @return trimmed string, or original reference if no trailing whitespace was found
     */
    private static String trimTrailing(String input)
    {
        if (input == null || input.length() == 0)
        {
            return input;
        }

        int len = input.length();

        while (len > 0 && isWhitespace(input.charAt(len - 1)))
        {
            len--;
        }

        return (len == input.length()) ? input : input.substring(0, len);
    }
}