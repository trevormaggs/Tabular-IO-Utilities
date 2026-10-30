package io.writer;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Locale;
import java.util.Objects;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import io.ExcelWritable;
import io.TabularDataException;
import io.style.AlternatingRowFactory;
import io.style.ColourStyle;

/**
 * Provides a partial implementation of the {@link ExcelWritable} contract for creating and writing
 * data to Excel workbooks using Apache POI.
 *
 * <p>
 * Subclasses are responsible for validating the workbook file extension and creating worksheets.
 * </p>
 *
 * <p>
 * <b>Change log:</b>
 * </p>
 *
 * <ul>
 * <li>2 November 2021 - Initial creation.</li>
 * <li>7 October 2026 - Simplified internal state access by declaring fields protected and removing
 * redundant accessors.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.7
 * @since 2 November 2021
 */
public abstract class AbstractWorkbook implements ExcelWritable
{
    private static final String DEFAULT_WORKSHEET_NAME = "Default";
    private String worksheetName = DEFAULT_WORKSHEET_NAME;
    private final Path filePath;
    private int rowIndex;
    private int columnIndex;
    private Row currentRow;
    private Cell currentCell;
    protected Workbook workbook;
    protected Sheet sheet;

    /**
     * Creates an instance configured to produce a new XLSX workbook using {@code out.xlsx} as the
     * default output file and {@code Default} as the default worksheet name.
     */
    public AbstractWorkbook()
    {
        filePath = Paths.get("out.xlsx");
    }

    /**
     * Creates an instance configured to produce a new Excel workbook using the specified output
     * file and the default worksheet name.
     *
     * <p>
     * The file extension is converted to lower case before the workbook type is validated.
     * </p>
     *
     * @param file
     *        the relative or absolute path of the output file
     */
    public AbstractWorkbook(String file)
    {
        Path tpath = null;

        if (file != null && !file.isEmpty())
        {
            Path fpath = Paths.get(file);
            Path fileName = fpath.getFileName();

            if (fileName != null)
            {
                String nameStr = fileName.toString();
                int extPos = nameStr.lastIndexOf('.');

                if (extPos >= 0)
                {
                    Path parent = fpath.getParent();
                    String normalised = nameStr.substring(0, extPos) + nameStr.substring(extPos).toLowerCase(Locale.ROOT);

                    tpath = (parent != null ? parent.resolve(normalised) : Paths.get(normalised));
                }
            }

            else
            {
                tpath = fpath;
            }
        }

        else
        {
            tpath = Paths.get("out.xlsx");
        }

        filePath = tpath;
    }

    /**
     * Creates an instance configured to produce a new Excel workbook with the specified worksheet
     * name.
     *
     * @param file
     *        the relative or absolute path to the output file
     * @param tabName
     *        the name of the worksheet to be created
     */
    public AbstractWorkbook(String file, String tabName)
    {
        this(file);
        worksheetName = tabName;
    }

    /**
     * Adds a new worksheet to the workbook and makes it the current worksheet.
     *
     * @param tabName
     *        the name of the worksheet to be created
     *
     * @throws TabularDataException
     *         if the worksheet cannot be created
     */
    @Override
    public void createNewWorksheet(String tabName) throws TabularDataException
    {
        currentRow = null;
        currentCell = null;
        rowIndex = 0;
        columnIndex = 0;
        worksheetName = tabName;

        addSheet();
    }

    /**
     * Gets the name of the current worksheet.
     *
     * @return the current worksheet name
     */
    @Override
    public String getWorksheetName()
    {
        return worksheetName;
    }

    /**
     * Adds an AutoFilter to the specified cell range of the current worksheet.
     *
     * @param firstRow
     *        zero-based index of the first row, inclusive
     * @param firstCol
     *        zero-based index of the first column, inclusive
     * @param lastRow
     *        zero-based index of the last row, inclusive
     * @param lastCol
     *        zero-based index of the last column, inclusive
     */
    @Override
    public void setAutoFilter(int firstRow, int firstCol, int lastRow, int lastCol)
    {
        sheet.setAutoFilter(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /**
     * Applies alternating row styles to the specified cell range using a predefined colour scheme.
     *
     * @param firstRow
     *        zero-based first row
     * @param firstCol
     *        zero-based first column
     * @param lastRow
     *        zero-based last row
     * @param lastCol
     *        zero-based last column
     * @param colour
     *        the {@link ColourStyle} defining the alternating row colours
     */
    @Override
    public void alternateRowStyles(int firstRow, int firstCol, int lastRow, int lastCol, ColourStyle colour)
    {
        AlternatingRowFactory factory = AlternatingRowFactory.newInstance(sheet, colour);
        factory.apply(firstRow, firstCol, lastRow, lastCol);
    }

    /**
     * Creates an {@link AlternatingRowFactory} associated with the current worksheet.
     *
     * @return a new {@link AlternatingRowFactory} for the current worksheet
     */
    @Override
    public AlternatingRowFactory getInstanceAlternatingRow()
    {
        return AlternatingRowFactory.newInstance(sheet);
    }

    /**
     * Adds the next row to the current worksheet.
     *
     * @return the zero-based position of the newly created row
     */
    @Override
    public int addRow()
    {
        columnIndex = 0;
        currentRow = sheet.createRow(rowIndex++);

        return getCurrentRowPosition();
    }

    /**
     * Gets the current row position number.
     *
     * @return the current zero-based row position index, or -1 if no row has been created yet
     */
    @Override
    public int getCurrentRowPosition()
    {
        return (currentRow == null ? -1 : currentRow.getRowNum());
    }

    /**
     * Inserts a text value into the next cell of the current row.
     *
     * @param word
     *        the text value to insert
     * @return the zero-based column index of the newly created cell
     */
    @Override
    public int insertCell(String word)
    {
        currentCell = createCell(word);
        return currentCell.getColumnIndex();
    }

    /**
     * Inserts a numeric value into the next cell of the current row.
     *
     * <p>
     * The value is converted to a {@code double} before being written to the workbook.
     * Consequently, numeric types requiring greater precision than {@code double} provides may lose
     * precision.
     * </p>
     *
     * @param <T>
     *        the type extending {@link Number}
     * @param num
     *        the numeric value to insert, or {@code null} for a blank cell
     * @return the zero-based column index of the newly created cell
     */
    @Override
    public <T extends Number> int insertCell(T num)
    {
        currentCell = createCell(num);
        return currentCell.getColumnIndex();
    }

    /**
     * Processes an array of headers and writes them onto the worksheet at the current row location.
     * Default styling includes bold 12-point text and top vertical alignment.
     *
     * @param headers
     *        an array comprising header strings
     *
     * @throws TabularDataException
     *         if any writing issue occurs
     */
    @Override
    public void writeHeader(String[] headers) throws TabularDataException
    {
        Objects.requireNonNull(headers, "Header values cannot be null.");

        if (sheet == null)
        {
            throw new TabularDataException("Unable to write to worksheet [" + worksheetName + "]. Ensure the workbook is open first.");
        }

        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);

        CellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setFont(font);

        addRow();

        for (String item : headers)
        {
            Cell cell = createCell(item);
            cell.setCellStyle(style);
        }
    }

    /**
     * Writes an array of regular entries to a new row in the current worksheet.
     *
     * @param entries
     *        an array of entries in string form
     *
     * @throws TabularDataException
     *         if there is a problem writing to the worksheet
     */
    @Override
    public void writeContent(String[] entries) throws TabularDataException
    {
        Objects.requireNonNull(entries, "Entry values cannot be null.");

        if (sheet == null)
        {
            throw new TabularDataException("Cannot write to worksheet [" + worksheetName + "]. Check the workbook is opened first.");
        }

        addRow();

        for (String item : entries)
        {
            insertCell(item);
        }
    }

    /**
     * Gets the path of the workbook output file.
     *
     * @return the path of the configured workbook output file
     */
    @Override
    public Path getFilePath()
    {
        return filePath;
    }

    /**
     * Creates a new in-memory workbook and adds its initial worksheet.
     *
     * @throws TabularDataException
     *         if the workbook type is unsupported or the worksheet cannot be created
     */
    @Override
    public void open() throws TabularDataException
    {
        if (workbook != null)
        {
            throw new TabularDataException("Workbook is already open.");
        }

        checkWorkbookExtensionName();

        String fileName = filePath.toString().toLowerCase();

        if (fileName.endsWith(".xls"))
        {
            workbook = new HSSFWorkbook();
        }

        else if (fileName.endsWith(".xlsx"))
        {
            workbook = new XSSFWorkbook();
        }

        try
        {
            addSheet();
        }

        catch (TabularDataException exc)
        {
            try
            {
                workbook.close();
            }

            catch (IOException closeExc)
            {
                exc.addSuppressed(closeExc);
            }

            finally
            {
                workbook = null;
                sheet = null;
            }

            throw exc;
        }
    }

    /**
     * Writes the in-memory workbook to the configured output file and closes the workbook resource.
     *
     * @throws TabularDataException
     *         if the workbook cannot be written to the output file
     */
    @Override
    public void close() throws TabularDataException
    {
        Workbook currentWorkbook = workbook;
        
        if (currentWorkbook != null)
        {
            TabularDataException exc0 = null;

            try (OutputStream fout = Files.newOutputStream(filePath))
            {
                currentWorkbook.write(fout);
            }

            catch (IOException exc1)
            {
                exc0 = new TabularDataException("Unable to write workbook to file [" + filePath + "]", exc1);
                throw exc0;
            }

            finally
            {
                try
                {
                    currentWorkbook.close();
                }

                catch (IOException exc2)
                {
                    TabularDataException exc3 = new TabularDataException("Unable to close workbook [" + filePath + "]", exc2);

                    if (exc0 != null)
                    {
                        exc0.addSuppressed(exc3);
                    }

                    else
                    {
                        throw exc3;
                    }
                }

                finally
                {
                    workbook = null;
                    sheet = null;
                    currentRow = null;
                    currentCell = null;
                    rowIndex = 0;
                    columnIndex = 0;
                }
            }
        }
    }

    /**
     * Returns the row currently being populated.
     *
     * @return the current {@link Row} instance, or {@code null} if no row has been created
     */
    public Row getCurrentRow()
    {
        return currentRow;
    }

    /**
     * Returns the cell currently being populated.
     *
     * @return the current {@link Cell} instance, or {@code null} if no cell has been created
     */
    public Cell getCurrentCell()
    {
        return currentCell;
    }

    /**
     * Helper method to create a string cell at the current row position and advance column index.
     *
     * @param word
     *        the string value to populate into the new cell
     * @return the created POI {@link Cell}
     */
    private Cell createCell(String word)
    {
        if (currentRow == null)
        {
            addRow();
        }

        Cell cell = currentRow.createCell(columnIndex++);
        cell.setCellValue(word);

        return cell;
    }

    /**
     * Helper method to create a numeric cell at the current row position and advance column index.
     *
     * @param <T>
     *        the type extending {@link Number}
     * @param num
     *        the numeric value to populate, or {@code null} for a blank cell
     * @return the created POI {@link Cell}
     */
    private <T extends Number> Cell createCell(T num)
    {
        if (currentRow == null)
        {
            addRow();
        }

        Cell cell = currentRow.createCell(columnIndex++);

        if (num == null)
        {
            cell.setBlank();
        }

        else
        {
            cell.setCellValue(num.doubleValue());
        }

        return cell;
    }

    /**
     * Defers validation of the workbook file extension to the subclass.
     *
     * @throws TabularDataException
     *         if the workbook file path does not have a supported extension
     */
    protected abstract void checkWorkbookExtensionName() throws TabularDataException;

    /**
     * Defers creation of the current worksheet to the subclass.
     *
     * @throws TabularDataException
     *         if the worksheet cannot be created
     */
    protected abstract void addSheet() throws TabularDataException;
}