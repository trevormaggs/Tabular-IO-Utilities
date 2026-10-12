package io.writer;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import io.TabularDataException;
import io.TabularDataException.ErrorCode;

/**
 * Extends the {@link WriteWorkbook} class to append a new worksheet tab to an existing Excel
 * spreadsheet instead of overwriting the workbook, or create a new workbook when the target file
 * does not exist.
 *
 * <p>
 * Supports both binary XLS and OOXML XLSX workbook formats.
 * </p>
 * 
 * <b>Change log:</b>
 * <ul>
 * <li>26 June 2018 - Initial creation.</li>
 * <li>7 October 2026 - Refactored package to io.writer, updated exception handling to
 * TabularDataException and migrated stream handling to NIO Files.</li>
 * </ul>
 * 
 * @author Trevor Maggs
 * @version 0.5
 * @since 26 June 2018
 */
public final class AppendWorkbook extends WriteWorkbook
{
    /**
     * Constructs an {@code AppendWorkbook} instance targeting an existing Excel spreadsheet and
     * specifying a new worksheet tab name.
     * 
     * @param file
     *        relative or absolute path to the workbook file
     * @param tabName
     *        name of the worksheet tab to add
     * 
     * @throws TabularDataException
     *         if the workbook path or worksheet name is invalid
     */
    public AppendWorkbook(String file, String tabName) throws TabularDataException
    {
        super(file, tabName);
    }

    /**
     * Opens the existing spreadsheet and adds a new worksheet tab. If the target file does not
     * exist, delegates to {@code super.open()} to create a new workbook.
     *
     * @throws TabularDataException
     *         if the workbook cannot be opened, is encrypted, or cannot be modified
     */
    @Override
    public void open() throws TabularDataException
    {
        try (InputStream in = Files.newInputStream(getFilePath()))
        {
            workbook = WorkbookFactory.create(in);
        }

        catch (NoSuchFileException exc)
        {
            super.open();
        }

        catch (EncryptedDocumentException exc)
        {
            throw new TabularDataException(ErrorCode.FILE_ACCESS_DENIED, "Unable to read the password-protected workbook [" + getFilePath() + "]. Check password.", exc);
        }

        catch (IOException exc)
        {
            throw new TabularDataException(ErrorCode.STREAM_READ_FAILED, "Problem while attempting to open spreadsheet [" + getFilePath() + "] for appending.", exc);
        }

        addSheet();
    }
}