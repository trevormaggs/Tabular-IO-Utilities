package io.reader;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import io.TabularDataException;
import io.TabularDataException.ErrorCode;

/**
 * Uses the Apache POI libraries to read an Excel spreadsheet or similar workbook and update its
 * contents. It supports both Binary XLS and XML-based XLSX formats.
 * 
 * <p>
 * <b>IMPORTANT:</b> Call {@link #save()} to persist updates back to disk. Once finished, call
 * {@link #close()} (or use inside a {@code try-with-resources} block) to release underlying POI
 * memory buffers and file handles.
 * </p>
 * 
 * <p>
 * <b>Change log:</b>
 * </p>
 * 
 * <ul>
 * <li>Trevor Maggs created on 8 November 2021</li>
 * <li>Trevor Maggs modified on 8 October 2026 - Refactored package location to io.reader,
 * modernised exception handling with TabularDataException, fixed font styling state, and decoupled
 * saving from close lifecycle.</li>
 * </ul>
 * 
 * @author Trevor Maggs
 * @version 0.4
 * @since 8 November 2021
 */
public final class UpdateWorkbook2 extends ReadWorkbook
{
    private final Path newFile;
    private CellStyle cellStyle;

    /**
     * Loads the specified source workbook for in-place updates to the same file.
     * 
     * @param file
     *        the Path object representing the workbook file
     * 
     * @throws TabularDataException
     *         if the workbook cannot be accessed or an error occurs while initialising the reader
     */
    public UpdateWorkbook2(Path file) throws TabularDataException
    {
        this(file, file);
    }

    /**
     * Loads the source workbook and writes updates out to a separate destination workbook file.
     * 
     * @param file1
     *        the Path object representing the source workbook file
     * @param file2
     *        the Path object representing the destination workbook file
     * 
     * @throws TabularDataException
     *         if the workbook cannot be accessed or an error occurs while initialising the reader
     */
    public UpdateWorkbook2(Path file1, Path file2) throws TabularDataException
    {
        super(file1);
        this.newFile = file2;
    }

    /**
     * Loads the source workbook and selects a specific worksheet tab for updating.
     * 
     * @param file1
     *        the Path object representing the source workbook file
     * @param file2
     *        the Path object representing the destination workbook file
     * @param tabParam
     *        either zero-based index in integer form or the tab name in string form to select the
     *        required worksheet
     * 
     * @throws TabularDataException
     *         if the workbook cannot be accessed or an error occurs while initialising the reader
     */
    public UpdateWorkbook2(Path file1, Path file2, Object tabParam) throws TabularDataException
    {
        super(file1, tabParam);
        this.newFile = file2;
    }

    /**
     * Sets a cell font colour for subsequent cell updates.
     *
     * <p>
     * See <a href=
     * "https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/IndexedColors.html">poi.apache.org/apidocs/org/apache/poi/ss/usermodel/IndexedColors.html</a>
     * </p>
     * 
     * @param colour
     *        the {@link IndexedColors} value to use for cell text
     * 
     * @throws TabularDataException
     *         if the workbook resource is not open
     */
    public void setFontColour(IndexedColors colour) throws TabularDataException
    {
        Workbook wb = workbook;

        if (wb == null)
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Workbook resource is not open. Call open() first.");
        }

        if (cellStyle == null)
        {
            cellStyle = wb.createCellStyle();
        }

        Font fontStyle = wb.createFont();
        fontStyle.setColor(colour.getIndex());
        cellStyle.setFont(fontStyle);
    }

    /**
     * Renames the selected worksheet.
     *
     * @param replace
     *        the new worksheet name
     *        
     * @throws TabularDataException
     *         if the worksheet is unknown        
     */
    public void renameWorksheet(String replace) throws TabularDataException
    {
        Sheet worksheet = sheet;

        if (worksheet == null)
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Undefined worksheet encountered.");
        }

        Workbook wb = worksheet.getWorkbook();
        wb.setSheetName(wb.getSheetIndex(worksheet), replace);
    }

    /**
     * Replaces matching content in the first row of the selected worksheet.
     *
     * @param text
     *        the content to find
     * @param replace
     *        the replacement string
     *
     * @return {@code true} if matching content was found and replaced; otherwise {@code false}
     *
     * @throws TabularDataException
     *         if no active worksheet is selected
     */
    public boolean replaceHeader(String text, String replace) throws TabularDataException
    {
        return replaceContent(text, replace, 0);
    }

    /**
     * Replaces the content in a cell located in the specified row position with the replacement
     * string provided.
     * 
     * @param text
     *        the content to find
     * @param replace
     *        the replacement string
     * @param rowIndex
     *        the zero-based row index in which to search
     * @return boolean true if the replacement was a success, otherwise false
     * 
     * @throws TabularDataException
     *         if no active worksheet is selected
     */
    public boolean replaceContent(String text, String replace, int rowIndex) throws TabularDataException
    {
        int col = findColumnIndexByContent(text);

        if (col > -1)
        {
            updateCellValue(rowIndex, col, replace);
            return true;
        }

        return false;
    }

    /**
     * Updates the content in a cell.
     *
     * @param row
     *        the zero-based row index
     * @param col
     *        the zero-based column index
     * @param text
     *        text to replace
     * 
     * @throws TabularDataException
     *         if no active worksheet is selected
     */
    public void updateCellValue(int row, int col, String text) throws TabularDataException
    {
        Sheet worksheet = sheet;

        if (worksheet == null)
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Undefined worksheet encountered.");
        }

        Row rowLine = worksheet.getRow(row);

        if (rowLine == null)
        {
            rowLine = worksheet.createRow(row);
        }

        Cell cell = rowLine.getCell(col);

        if (cell == null)
        {
            cell = rowLine.createCell(col);
        }

        cell.setCellValue(text);

        if (cellStyle != null)
        {
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * Updates the cell numeric content.
     *
     * @param <T>
     *        expected Number type parameter
     * @param row
     *        the row number
     * @param col
     *        the column number
     * @param value
     *        the numerical entry to be inserted in the cell
     * 
     * @throws TabularDataException
     *         if no active worksheet is selected
     */
    public <T extends Number> void updateCellValue(int row, int col, T value) throws TabularDataException
    {
        Sheet worksheet = sheet;

        if (worksheet == null)
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Undefined worksheet encountered.");
        }

        Row rowLine = worksheet.getRow(row);

        if (rowLine == null)
        {
            rowLine = worksheet.createRow(row);
        }

        Cell cell = rowLine.getCell(col);

        if (cell == null)
        {
            cell = rowLine.createCell(col);
        }

        cell.setCellValue(value.doubleValue());

        if (cellStyle != null)
        {
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * Opens and acquires the resource pointing to the specified workbook.
     * 
     * @throws TabularDataException
     *         if the file does not exist, is password-protected, or an I/O error occurs
     */
    @Override
    public void open() throws TabularDataException
    {
        Path fpath = getFilePath();

        if (!Files.isRegularFile(fpath))
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "File [" + fpath + "] does not exist or is not a regular file.");
        }

        try
        {
            workbook = WorkbookFactory.create(fpath.toFile());
            selectSheet();
        }

        catch (FileNotFoundException exc)
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "File [" + fpath + "] does not exist.", exc);
        }

        catch (EncryptedDocumentException exc)
        {
            throw new TabularDataException(ErrorCode.FILE_ACCESS_DENIED, "Unable to read the password-protected workbook [" + fpath + "]. Check the password.", exc);
        }

        catch (IOException exc)
        {
            throw new TabularDataException(ErrorCode.STREAM_READ_FAILED, "An error occurred while reading data in file [" + fpath + "].", exc);
        }
    }

    /**
     * Writes updated workbook contents to the target destination file using a temporary file before
     * replacing the existing destination file.
     *
     * <p>
     * If an error occurs while writing the temporary file or replacing the destination file, the
     * temporary file is removed where possible.
     * </p>
     *
     * @throws TabularDataException
     *         if an I/O error occurs while persisting data to disk
     */
    public void save() throws TabularDataException
    {
        Workbook wb = workbook;

        if (wb == null)
        {
            throw new TabularDataException(ErrorCode.STREAM_WRITE_FAILED, "Cannot save workbook. Call open() first.");
        }

        Path tempFile = null;

        try
        {
            tempFile = Files.createTempFile("tabular_update_", ".tmp");

            try (OutputStream fout = Files.newOutputStream(tempFile))
            {
                wb.write(fout);
            }

            Files.move(tempFile, newFile, StandardCopyOption.REPLACE_EXISTING);
        }

        catch (IOException exc)
        {
            if (tempFile != null)
            {
                try
                {
                    Files.deleteIfExists(tempFile);
                }

                catch (IOException exc2)
                {
                    // Preserve the cleanup failure without replacing the original exception
                    exc.addSuppressed(exc2);
                }
            }

            throw new TabularDataException(ErrorCode.STREAM_WRITE_FAILED, "Unable to write data to worksheet [" + getWorksheetName() + "] in file [" + newFile + "]", exc);
        }
    }
}