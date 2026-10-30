package io.writer;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardOpenOption;
import io.TabularDataException;
import io.TabularDataException.ErrorCode;

/**
 * Writes additional rows to an existing CSV file without overwriting its contents.
 * 
 * <p>
 * <b>Change log:</b>
 * </p>
 * 
 * <ul>
 * <li>Released by Trevor Maggs on 10 August 2020</li>
 * </ul>
 * 
 * @author Trevor Maggs
 * @version 0.2
 * @since 10 August 2020
 */
public class AppendCSV extends WriteCSV
{
    /**
     * Constructs an {@code AppendCSV} writer targeting the specified file path.
     * 
     * @param file
     *        relative or absolute file path
     */
    public AppendCSV(String file)
    {
        super(file);
    }

    /**
     * Opens the target CSV file and initialises the writer in append mode.
     * 
     * @throws TabularDataException
     *         if an I/O error occurs while opening the file
     */
    @Override
    public void open() throws TabularDataException
    {
        if (!filePath.toString().endsWith(".csv"))
        {
            throw new TabularDataException("The file [" + filePath + "] does not appear to be a valid CSV file.");
        }

        if (csvWriter != null)
        {
            throw new TabularDataException("CSV file is already open.");
        }

        try
        {
            csvWriter = Files.newBufferedWriter(filePath, StandardCharsets.UTF_8, StandardOpenOption.CREATE, StandardOpenOption.APPEND);
        }

        catch (IOException exc)
        {
            throw new TabularDataException(ErrorCode.FILE_ACCESS_DENIED, "Cannot open CSV file [" + filePath + "] for appending. Check file permissions.", exc);
        }
    }
}