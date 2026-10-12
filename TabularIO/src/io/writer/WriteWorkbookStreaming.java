package io.writer;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import io.TabularDataException;

/**
 * Streaming Excel workbook writer extending {@link AbstractWorkbook} using Apache POI's
 * {@link SXSSFWorkbook}.
 * 
 * <p>
 * {@code SXSSFWorkbook} is an API-compatible streaming extension of {@code XSSF} used when very
 * large spreadsheets must be produced with limited heap space. Low memory consumption is achieved
 * by retaining a sliding window of active rows in memory while flushing older rows to temporary
 * files on disk.
 * </p>
 * 
 * <p>
 * <b>Change log:</b>
 * </p>
 * <ul>
 * <li>26 October 2021 - Initial implementation.</li>
 * <li>6 October 2026 - Updated resource cleanup, NIO output streams, and open-time extension
 * checks.</li>
 * </ul>
 * 
 * @author Trevor Maggs
 * @version 0.3
 * @since 26 October 2021
 */
public class WriteWorkbookStreaming extends AbstractWorkbook
{
    private boolean sharedStringsTable;
    private SXSSFWorkbook sxssfWorkbook;

    /**
     * Instantiates a new streaming Excel workbook using {@code out.xlsx} as the default file path
     * and {@code Default} as the initial worksheet tab name.
     */
    public WriteWorkbookStreaming()
    {
        super();
    }

    /**
     * Instantiates a new streaming Excel workbook for the specified target file path using the
     * default worksheet tab name.
     * 
     * @param file
     *        the relative or absolute path to the target XLSX file
     * 
     * @throws TabularDataException
     *         if the specified file path cannot be resolved
     */
    public WriteWorkbookStreaming(String file) throws TabularDataException
    {
        super(file);
    }

    /**
     * Instantiates a new streaming Excel workbook for the specified target file path and worksheet
     * name. Shared strings table is disabled by default for memory efficiency.
     * 
     * @param file
     *        the relative or absolute path to the target XLSX file
     * @param tabName
     *        the name of the initial worksheet tab to create
     * 
     * @throws TabularDataException
     *         if the specified file path cannot be resolved
     */
    public WriteWorkbookStreaming(String file, String tabName) throws TabularDataException
    {
        this(file, tabName, false);
    }

    /**
     * Instantiates a new streaming Excel workbook with options for enabling the shared strings
     * table.
     * 
     * <p>
     * <b>Note:</b> Enabling the shared strings table may reduce resultant file size, but requires
     * Apache POI to keep all distinct string literals in memory during processing. For extremely
     * large datasets, keeping this option {@code false} is recommended to prevent high heap usage
     * or {@link OutOfMemoryError}.
     * </p>
     * 
     * @param file
     *        the relative or absolute path to the target XLSX file
     * @param tabName
     *        the name of the initial worksheet tab to create
     * @param useSharedStringsTable
     *        {@code true} to enable the shared strings table; {@code false} to write strings inline
     * 
     * @throws TabularDataException
     *         if the specified file path cannot be resolved
     */
    public WriteWorkbookStreaming(String file, String tabName, boolean useSharedStringsTable) throws TabularDataException
    {
        super(file, tabName);
        sharedStringsTable = useSharedStringsTable;
    }

    /**
     * Validates the output file extension, initialises the underlying {@link SXSSFWorkbook}
     * resource, and creates the default worksheet.
     * 
     * <p>
     * If this workbook instance is already open, calling this method safely disposes and closes
     * existing underlying streaming resources prior to re-initialisation.
     * </p>
     * 
     * @throws TabularDataException
     *         if the target file path does not end with {@code .xlsx} or if worksheet creation
     *         fails
     */
    @Override
    public void open() throws TabularDataException
    {
        checkWorkbookExtensionName();

        if (sxssfWorkbook != null)
        {
            close();
        }

        if (sharedStringsTable)
        {
            sxssfWorkbook = new SXSSFWorkbook(null, SXSSFWorkbook.DEFAULT_WINDOW_SIZE, false, true);
        }

        else
        {
            sxssfWorkbook = new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
        }

        workbook = sxssfWorkbook.getXSSFWorkbook();

        addSheet();
    }

    /**
     * Writes the streaming workbook contents to the target file on disk, cleans up temporary row
     * disk files, and closes all open workbook resources.
     * 
     * @throws TabularDataException
     *         if an I/O error occurs while opening or writing data to the output file stream
     */
    @Override
    public void close() throws TabularDataException
    {
        if (sxssfWorkbook == null)
        {
            return;
        }

        try (OutputStream fout = Files.newOutputStream(getFilePath()))
        {
            sxssfWorkbook.write(fout);
        }

        catch (IOException exc)
        {
            throw new TabularDataException("Unable to write data to worksheet [" + getWorksheetName() + "] in file [" + getFilePath() + "]", exc);
        }

        finally
        {
            try
            {
                sxssfWorkbook.dispose();
            }

            finally
            {
                try
                {
                    sxssfWorkbook.close();
                }

                catch (IOException exc)
                {
                    // Ignore close failure after write attempt
                }

                finally
                {
                    sxssfWorkbook = null;
                }
            }
        }
    }

    /**
     * Validates that the specified file path ends with a supported XML-formatted XLSX extension.
     * 
     * @throws TabularDataException
     *         if the specified file extension is not {@code .xlsx}
     */
    @Override
    protected void checkWorkbookExtensionName() throws TabularDataException
    {
        String file = getFilePath().toString();

        if (!file.endsWith(".xlsx"))
        {
            throw new TabularDataException("The file [" + file + "] does not appear to be in a valid XLSX format.");
        }
    }

    /**
     * Creates a new worksheet tab on the active streaming workbook using the current worksheet
     * name.
     * 
     * @throws TabularDataException
     *         if the underlying {@link SXSSFWorkbook} instance has not been opened yet
     */
    @Override
    protected void addSheet() throws TabularDataException
    {
        if (sxssfWorkbook == null)
        {
            throw new TabularDataException("Cannot add a new worksheet to spreadsheet [" + getFilePath() + "]. You must open the workbook first.");
        }

        sheet = sxssfWorkbook.createSheet(getWorksheetName());
    }
}