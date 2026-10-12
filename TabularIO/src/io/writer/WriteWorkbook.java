package io.writer;

import io.ExcelWritable;
import io.TabularDataException;

/**
 * Provides an implementation of {@link ExcelWritable} for creating and writing entries to an Excel
 * workbook using the Apache POI library. The workbook can be created in either the binary XLS
 * format or the XML-based XLSX format.
 *
 * <p>
 * <b>Change log:</b>
 * </p>
 *
 * <ul>
 * <li>Created on 2 November 2021</li>
 * <li>Updated to inherit the {@link ExcelWritable} implementation via {@link AbstractWorkbook}</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.6
 * @since 2 November 2021
 */
public class WriteWorkbook extends AbstractWorkbook
{
    /**
     * Creates a new Excel workbook using the default workbook file and worksheet name.
     */
    public WriteWorkbook()
    {
        super();
    }

    /**
     * The constructor is used to create a new Excel spreadsheet. The worksheet to be created willc
     * be given a default name.
     *
     * @param file
     *        the relative or absolute path of the output file
     *
     * @throws TabularDataException
     *         if the specified file does not have a valid XLS or XLSX extension
     */
    public WriteWorkbook(String file) throws TabularDataException
    {
        super(file);
    }

    /**
     * The constructor is used to create a new Excel spreadsheet with the new worksheet to be
     * labelled with the specified sheet tab name.
     *
     * @param file
     *        the relative or absolute path to the file
     * @param tabName
     *        the name of the worksheet to be created
     *
     * @throws TabularDataException
     *         if the type of the specified spreadsheet is not a valid XLS or XLSX format
     */
    public WriteWorkbook(String file, String tabName) throws TabularDataException
    {
        super(file, tabName);
    }

    /**
     * Validates that the specified file has an XLS or XLSX extension.
     *
     * @throws TabularDataException
     *         if the specified file does not have an XLS or XLSX extension
     */
    @Override
    protected void checkWorkbookExtensionName() throws TabularDataException
    {
        String file = getFilePath().toString();

        if (!(file.endsWith(".xls") || file.endsWith(".xlsx")))
        {
            throw new TabularDataException("The file [" + file + "] does not appear to be a valid Excel spreadsheet.");
        }
    }

    /**
     * Creates and adds a worksheet using the current worksheet name.
     *
     * <p>
     * If a worksheet with that name already exists, a numeric suffix is appended incrementally
     * until a unique name is found.
     * </p>
     *
     * @throws TabularDataException
     *         if the workbook has not been opened
     */
    @Override
    protected void addSheet() throws TabularDataException
    {
        if (workbook != null)
        {
            int suffix = 0;
            String sheetName = getWorksheetName();

            while (workbook.getSheet(sheetName) != null)
            {
                sheetName = String.format("%s_%d", getWorksheetName(), ++suffix);
            }

            sheet = workbook.createSheet(sheetName);
        }

        else
        {
            throw new TabularDataException("Cannot add a new worksheet to spreadsheet [" + getFilePath() + "]. You must open the workbook first.");
        }
    }
}