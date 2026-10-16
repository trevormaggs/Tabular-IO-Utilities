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
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;
import io.Readable;
import io.TabularDataException;
import io.TabularDataException.ErrorCode;

/**
 * Reads CSV files into memory using a quote-aware state-machine parser that supports quoted fields,
 * escaped quotes, and multiline records.
 *
 * <p>
 * Implements {@link Readable} and {@link Iterable} for uniform table processing. The base class
 * automatically sanitises internal newlines during row processing. Subclasses can override
 * {@link #processHeader(String[])} or extend {@link #processRow(String[], int)} to apply custom
 * validation, schema checking, or domain-specific data transformations.
 * </p>
 *
 * <b>Change log:</b>
 *
 * <ul>
 * <li>7 February 2025 - Initial creation as AbstractCSVReader.</li>
 * <li>8 October 2026 - Re-factored to concrete ReadCSV in io.reader, added RFC 4180 multi-line
 * record parsing, and integrated TabularDataException.</li>
 * <li>10 October 2026 - Updated matrix column sizing to evaluate maximum row width across the whole
 * document.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.6
 * @since 7 February 2025
 */
public class ReadCSV implements Readable
{
    private final Path filePath;
    private final List<String[]> dataset;
    private int lineNumber;
    private int maxColumnSize;

    /**
     * Constructs a {@code ReadCSV} instance targeting the specified file path.
     *
     * @param csvFile
     *        path to the target CSV file
     *
     * @throws TabularDataException
     *         if the file path is null or does not point to a regular file
     */
    public ReadCSV(Path csvFile) throws TabularDataException
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
    }

    /**
     * Constructs a {@code ReadCSV} instance for the specified CSV file path.
     *
     * @param csvFile
     *        path to the target CSV file
     *
     * @throws InvalidPathException
     *         if the specified path string is invalid
     * @throws TabularDataException
     *         if the path is null or does not point to a regular file
     */
    public ReadCSV(String csvFile) throws TabularDataException
    {
        this(csvFile == null ? null : Paths.get(csvFile));
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
     * Opens and parses the CSV file contents into memory, tokenising multi-line records andapplying
     * {@link #processRow(String[], int)} sanitisation to populated rows.
     *
     * @throws TabularDataException
     *         if the file cannot be accessed, does not exist, or an I/O error occurs
     */
    @Override
    public void open() throws TabularDataException
    {
        dataset.clear();
        maxColumnSize = 0;

        try (BufferedReader br = Files.newBufferedReader(filePath, StandardCharsets.UTF_8))
        {
            lineNumber = 0;
            String[] record;

            while ((record = parseDataCSV(br)) != null)
            {
                if (lineNumber == 0)
                {
                    processHeader(record);
                }

                else
                {
                    processRow(record, lineNumber);
                }

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
     * Optional hook method executed when the header row (row 0) is parsed.
     *
     * @param headers
     *        array of header tokens
     */
    protected void processHeader(String[] headers)
    {
        // Default implementation does nothing. Subclasses can override if needed
    }

    /**
     * Executes post-processing on each parsed content row greater than zero (first row) position.
     *
     * <p>
     * The base implementation flattens embedded line breaks into single spaces and trims cell
     * values. Subclasses overriding this method should invoke
     * {@code super.processRow(rowTokens, lineNumber)} if they wish to retain default string
     * sanitisation prior to executing custom validation.
     * </p>
     *
     * @param rowTokens
     *        array of raw cell tokens for the current record
     * @param lineNumber
     *        1-based data row index (where header is row 0)
     */
    protected void processRow(String[] rowTokens, int lineNumber)
    {
        for (int i = 0; i < rowTokens.length; i++)
        {
            if (rowTokens[i] != null)
            {
                rowTokens[i] = rowTokens[i].replaceAll("[\\r\\n]+", " ").trim();
            }
        }
    }

    /**
     * Parses a single multi-line CSV record from a BufferedReader according to RFC 4180. Handles
     * quotes, escaped quotes (""), embedded newlines, custom whitespace (including NBSP), and
     * leading/trailing unquoted trimming.
     *
     * @param br
     *        BufferedReader stream positioned at the record start
     * @return String[] representing tokenised field values, or null if EOF reached
     * 
     * @throws IOException
     *         if an unterminated double-quoted field is encountered
     */
    private String[] parseDataCSV(BufferedReader br) throws IOException
    {
        String line = br.readLine();

        if (line != null)
        {
            List<String> tokens = new ArrayList<>();
            StringBuilder sb = new StringBuilder();
            boolean inQuotes = false;
            boolean fieldStart = true;

            while (line != null)
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
                line = br.readLine();
            }

            // Reached end-of-file inside quotes
            throw new IOException("Malformed CSV: Unterminated double-quoted field encountered.");
        }

        return null; // Stream end reached
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
}