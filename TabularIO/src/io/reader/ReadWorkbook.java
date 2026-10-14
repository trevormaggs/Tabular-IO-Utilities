package io.reader;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.Locale;
import java.util.NoSuchElementException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import io.ExcelReadable;
import io.TabularDataException;
import io.TabularDataException.ErrorCode;

/**
 * Reads Excel workbooks in binary XLS and XML-based XLSX formats using Apache POI.
 *
 * <p>
 * Implements {@link ExcelReadable} to provide worksheet selection, cell access, and row iteration.
 * The reader must be opened using {@link #open()} before accessing worksheet contents and should be
 * closed after use to release underlying resources. A {@code try-with-resources} statement is
 * recommended.
 * </p>
 *
 * <b>Change log:</b>
 * <ul>
 * <li>4 November 2021 - Initial creation.</li>
 * <li>7 October 2026 - Refactored package location to io.reader, implemented ExcelReadable and
 * migrated exception handling to TabularDataException.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.4
 * @since 4 November 2021
 */
public class ReadWorkbook implements ExcelReadable
{
    private final Path filePath;
    private Object tab;
    private int sheetCount;
    private SheetDimension matrixsize;
    protected Workbook workbook;
    protected Sheet sheet;

    private final DataFormatter dataFormatter = new DataFormatter();

    /**
     * Constructs a {@code ReadWorkbook} instance targeting the specified spreadsheet file.
     *
     * @param fpath
     *        path to the workbook file
     * @param tabParam
     *        zero-based worksheet index or worksheet name
     * @throws TabularDataException
     *         if the path is null, the file does not exist, or its extension is not a supported
     *         Excel format (.xls or .xlsx)
     */
    public ReadWorkbook(Path fpath, Object tabParam) throws TabularDataException
    {
        if (fpath == null)
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "Target path cannot be null.");
        }

        if (Files.notExists(fpath))
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "File [" + fpath + "] does not exist.");
        }

        String fname = fpath.toString().toLowerCase(Locale.ROOT);

        if (!fname.endsWith(".xls") && !fname.endsWith(".xlsx"))
        {
            throw new TabularDataException(ErrorCode.UNSUPPORTED_FORMAT, "File [" + fpath + "] does not appear to be a valid Excel spreadsheet format.");
        }

        this.filePath = fpath;
        this.tab = tabParam;
    }

    /**
     * Constructs a {@code ReadWorkbook} instance targeting the specified spreadsheet file,
     * selecting the default worksheet, typically the first tab.
     *
     * @param fpath
     *        path to the workbook file
     *
     * @throws TabularDataException
     *         if the path is null, the file does not exist, or its extension is not a supported
     *         Excel format (.xls or .xlsx)
     */
    public ReadWorkbook(Path fpath) throws TabularDataException
    {
        this(fpath, null);
    }

    /**
     * Constructs a {@code ReadWorkbook} instance targeting the specified spreadsheet file,
     * selecting the default worksheet, typically the first tab.
     *
     * @param file
     *        path to the workbook file in string form
     *
     * @throws InvalidPathException
     *         if the specified path string is invalid
     * @throws TabularDataException
     *         if the path is null, the file does not exist, or its extension is not a supported
     *         Excel format (.xls or .xlsx)
     */
    public ReadWorkbook(String file) throws TabularDataException
    {
        this(Paths.get(file), null);
    }

    /**
     * Constructs a {@code ReadWorkbook} instance targeting the specified spreadsheet file with the
     * given worksheet name selected.
     *
     * @param file
     *        path to the workbook file in string form
     * @param tabParam
     *        zero-based worksheet index or worksheet name
     * 
     * @throws InvalidPathException
     *         if the specified path string is invalid
     * @throws TabularDataException
     *         if the path is null, the file does not exist, or its extension is not a supported
     *         Excel format (.xls or .xlsx)
     */
    public ReadWorkbook(String file, Object tabParam) throws TabularDataException
    {
        this(Paths.get(file), tabParam);
    }

    /**
     * Gets the path of the workbook file.
     *
     * @return the path of the configured workbook file
     */
    @Override
    public Path getFilePath()
    {
        return filePath;
    }

    /**
     * Opens the workbook, records the number of worksheets, and selects the configured worksheet
     * when one is specified.
     *
     * @throws TabularDataException
     *         if the workbook cannot be opened or read
     */
    @Override
    public void open() throws TabularDataException
    {
        try
        {
            workbook = WorkbookFactory.create(filePath.toFile());
            sheetCount = workbook.getNumberOfSheets();
            selectSheet();
        }

        catch (FileNotFoundException exc)
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "File [" + filePath + "] does not exist.", exc);
        }

        catch (EncryptedDocumentException exc)
        {
            throw new TabularDataException(ErrorCode.FILE_ACCESS_DENIED, "Unable to read the password-protected workbook [" + filePath + "]. Check password.", exc);
        }

        catch (IOException exc)
        {
            throw new TabularDataException(ErrorCode.STREAM_READ_FAILED, "An error occurred while reading data in file [" + filePath + "].", exc);
        }
    }

    /**
     * Closes the underlying workbook resource and releases memory buffers.
     *
     * @throws TabularDataException
     *         if an I/O error occurs while closing the workbook
     */
    @Override
    public void close() throws TabularDataException
    {
        if (workbook != null)
        {
            try
            {
                workbook.close();
            }

            catch (IOException exc)
            {
                throw new TabularDataException(ErrorCode.STREAM_READ_FAILED, "Problem closing workbook resource for file [" + filePath + "].", exc);
            }

            finally
            {
                workbook = null;
                sheet = null;
                matrixsize = null;
                sheetCount = 0;
            }
        }
    }

    /**
     * Gets the row count of the active worksheet, including any empty rows within its used range.
     *
     * @return row count, or 0 if no worksheet is selected
     */
    @Override
    public int getRowCount()
    {
        return (matrixsize != null ? matrixsize.getRows() : 0);
    }

    /**
     * Gets the column count of the active worksheet's used range.
     *
     * @return column count, or 0 if no worksheet is selected
     */
    @Override
    public int getColumnCount()
    {
        return (matrixsize != null ? matrixsize.getColumns() : 0);
    }

    /**
     * Finds the column index whose cell content exactly matches the specified text, ignoring case.
     *
     * @param entry
     *        literal cell content to locate
     * @param rowIndex
     *        zero-based row index to search
     * @return zero-based column index, or {@code -1} if not found
     */
    @Override
    public int findColumnIndexByContent(String entry, int rowIndex)
    {
        return findColumnIndexRegex(String.format("^%s$", entry), rowIndex);
    }

    /**
     * Finds the column index whose cell content matches the specified regular expression, ignoring
     * case.
     *
     * @param regex
     *        regular expression pattern to evaluate against cell content
     * @param rowIndex
     *        zero-based row index to search
     * @return zero-based column index, or {@code -1} if no match is found, regex is empty/null, or
     *         no worksheet is active
     */
    @Override
    public int findColumnIndexRegex(String regex, int rowIndex)
    {
        if (sheet != null && regex != null && !regex.isEmpty())
        {
            if (rowIndex >= 0 && rowIndex < getRowCount())
            {
                Row row = sheet.getRow(rowIndex);

                if (row != null)
                {
                    Pattern pattern = null;

                    try
                    {
                        pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
                    }

                    catch (PatternSyntaxException exc)
                    {
                        // Fall back to exact literal matching if user supplied invalid regex
                        pattern = Pattern.compile(Pattern.quote(regex), Pattern.CASE_INSENSITIVE);
                    }

                    for (int j = 0; j < getColumnCount(); j++)
                    {
                        Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        Matcher matcher = pattern.matcher(formatCellToString(cell));

                        if (matcher.matches())
                        {
                            return j;
                        }
                    }
                }
            }
        }

        return -1;
    }

    /**
     * Returns an iterator across each row of the active worksheet as arrays of cell string values.
     *
     * @return row iterator
     */
    @Override
    public Iterator<String[]> iterator()
    {
        return new Iterator<String[]>()
        {
            private int currentIndex = 0;
            private final int endIndex = getRowCount();
            private final int columnCount = getColumnCount();

            @Override
            public boolean hasNext()
            {
                return (sheet != null && currentIndex >= 0 && currentIndex < endIndex);
            }

            @Override
            public String[] next()
            {
                if (!hasNext())
                {
                    throw new NoSuchElementException();
                }

                Row row = sheet.getRow(currentIndex++);
                String[] colArr = new String[columnCount];

                for (int j = 0; j < colArr.length; j++)
                {
                    colArr[j] = (row == null ? "" : formatCellToString(row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)));
                }

                return colArr;
            }
        };
    }

    /**
     * Returns the name of the currently active worksheet tab.
     *
     * @return worksheet tab name, or {@code null} if no worksheet is currently active
     */
    @Override
    public String getWorksheetName()
    {
        if (sheet != null)
        {
            return sheet.getSheetName();
        }

        return null;
    }

    /**
     * Returns the total number of worksheets in the workbook.
     *
     * @return total sheet count
     */
    @Override
    public int getSheetCount()
    {
        return sheetCount;
    }

    /**
     * Selects a worksheet by its zero-based index in the open workbook.
     *
     * @param tabIndex
     *        zero-based worksheet index
     * @throws TabularDataException
     *         if the workbook is not open or the index is invalid
     */
    @Override
    public void readWorksheet(int tabIndex) throws TabularDataException
    {
        if (workbook == null)
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Workbook [" + filePath + "] has not been opened.");
        }

        if (tabIndex < 0 || tabIndex >= sheetCount)
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Worksheet index must be between 0 and " + (sheetCount - 1) + ". Found [" + tabIndex + "].");
        }

        sheet = workbook.getSheetAt(tabIndex);
        measureDimensionSize();
    }

    /**
     * Selects a worksheet by name in the open workbook.
     *
     * @param tabName
     *        name of the worksheet
     * @throws TabularDataException
     *         if the workbook is not open or the worksheet cannot be found
     */
    @Override
    public void readWorksheet(String tabName) throws TabularDataException
    {
        if (workbook == null)
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Workbook [" + filePath + "] has not been opened.");
        }

        Sheet selectedSheet = tabName == null ? null : workbook.getSheet(tabName);

        if (selectedSheet == null)
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Worksheet [" + (tabName == null || tabName.trim().isEmpty() ? "BLANK" : tabName) + "] of file [" + filePath + "] is unknown.");
        }

        sheet = selectedSheet;
        measureDimensionSize();
    }

    /**
     * Returns the formatted string value of a cell at the specified row and column coordinates.
     *
     * @param row
     *        zero-based row index
     * @param col
     *        zero-based column index
     * @return cell value formatted as string, or an empty string if the cell is blank, out of
     *         bounds, or if no worksheet is active
     */
    public String getCellValue(int row, int col)
    {
        if (sheet != null)
        {
            if (row >= 0 && row < getRowCount() && col >= 0 && col < getColumnCount())
            {
                Row r = sheet.getRow(row);

                if (r != null)
                {
                    Cell c = r.getCell(col);

                    if (c != null)
                    {
                        return formatCellToString(c);
                    }
                }
            }
        }

        return "";
    }

    /**
     * Selects the initial worksheet and records sheet count.
     * 
     * @throws TabularDataException
     *         if the selected worksheet is unknown or undefined
     */
    protected void selectSheet() throws TabularDataException
    {
        if (tab instanceof Integer)
        {
            readWorksheet(((Integer) tab).intValue());
        }

        else if (tab instanceof String)
        {
            readWorksheet((String) tab);
        }

        else if (tab == null && sheetCount > 0)
        {
            readWorksheet(0);
        }

        else
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Worksheet [" + tab + "] is unknown or undefined.");
        }
    }

    /**
     * Formats a cell value as a string. Formula cells are evaluated to return their calculated
     * results.
     *
     * @param cell
     *        Apache POI cell to format
     * @return formatted cell value, or an empty string if the cell is null
     */
    private String formatCellToString(Cell cell)
    {
        if (cell.getCellType() == CellType.FORMULA)
        {
            try
            {
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                CellValue value = evaluator.evaluate(cell);

                if (value != null)
                {
                    switch (value.getCellType())
                    {
                        case NUMERIC:
                        case STRING:
                        case BOOLEAN:
                            return dataFormatter.formatCellValue(cell, evaluator).trim();

                        case ERROR:
                            return "ERROR: " + value.getErrorValue();

                        default:
                            return "";
                    }
                }
            }
            
            catch (Exception exc)
            {
                // Fallback to raw formula string if evaluation fails
            }

            return cell.getCellFormula().trim();
        }

        return dataFormatter.formatCellValue(cell).trim();
    }

    /**
     * Computes maximum active row and column boundaries for the active worksheet.
     */
    private void measureDimensionSize()
    {
        int x = 0;
        int y = 0;

        if (sheet != null && sheet.getPhysicalNumberOfRows() > 0)
        {
            x = sheet.getLastRowNum() + 1;

            for (Row row : sheet)
            {
                y = Math.max(y, row.getLastCellNum());
            }
        }

        matrixsize = new SheetDimension(x, y);
    }

    /**
     * Internal holder for row and column boundary counts.
     */
    private static final class SheetDimension
    {
        private final int rows;
        private final int columns;

        public SheetDimension(int rows, int columns)
        {
            this.rows = rows;
            this.columns = columns;
        }

        public int getRows()
        {
            return rows;
        }

        public int getColumns()
        {
            return columns;
        }
    }
}