package io.reader;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;
import javax.xml.namespace.QName;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import io.ExcelReadable;
import io.TabularDataException;
import io.TabularDataException.ErrorCode;

/**
 * Reads Excel XLSX workbooks using Apache POI’s event-based API to reduce memory consumption when
 * processing large worksheets.
 *
 * <p>
 * Implements {@link ExcelReadable} to provide tabular access to worksheet data without constructing
 * a complete in-memory worksheet model, such as a memory-intensive DOM model and other
 * alternatives.
 * </p>
 *
 * <p>
 * Worksheet XML is parsed using a SAX parser, allowing cell data to be processed incrementally
 * rather than loading the entire worksheet into a DOM tree. This approach can reduce memory
 * consumption, although actual usage depends on the workbook contents and the data retained by the
 * reader.
 * </p>
 *
 * <p>
 * Shared strings are accessed through Apache POI’s {@link ReadOnlySharedStringsTable}, which
 * provides read-only access to shared-string data without requiring the standard mutable shared
 * strings table.
 * </p>
 *
 * <p>
 * Cell values are exposed as strings. Loading the workbook’s style table is optional and enables
 * style-dependent interpretation of values, such as dates and specially formatted numbers. The
 * style table can be excluded when this interpretation is not required.
 * </p>
 *
 * <p>
 * <b>Change history:</b>
 * </p>
 * <ul>
 * <li>10 December 2017 — Initial implementation by Trevor Maggs.</li>
 * <li>9 October 2026 — Reorganised under {@code io.reader}, introduced
 * {@link TabularDataException}, reviewed stream handling, and standardised method ordering.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.5
 * @since 10 December 2017
 */
public class EventExcelReader implements ExcelReadable
{
    private final Path filePath;
    private final boolean useStyleTable;
    private final Object tabParam;
    private ExcelHandler handler;
    private String currentSheet;

    /**
     * Creates a reader for the specified workbook, selecting its first worksheet.
     *
     * @param fpath
     *        path to the XLSX workbook
     *
     * @throws TabularDataException
     *         if the path is invalid or the file is not an XLSX file
     */
    public EventExcelReader(Path fpath) throws TabularDataException
    {
        this(fpath, false, null);
    }

    /**
     * Creates a reader for the specified workbook, selecting its first worksheet.
     *
     * @param fpath
     *        path to the XLSX workbook
     * @param useStyleTable
     *        whether to load the workbook's style table
     *
     * @throws TabularDataException
     *         if the path is invalid or the file is not an XLSX file
     */
    public EventExcelReader(Path fpath, boolean useStyleTable) throws TabularDataException
    {
        this(fpath, useStyleTable, null);
    }

    /**
     * Creates a reader for the specified worksheet.
     *
     * @param fpath
     *        path to the XLSX workbook
     * @param useStyleTable
     *        whether to load the workbook's style table
     * @param tabParam
     *        zero-based worksheet index or worksheet name or {@code null} selects the first
     *        worksheet
     *
     * @throws TabularDataException
     *         if the path is invalid or the file is not an XLSX file
     */
    public EventExcelReader(Path fpath, boolean useStyleTable, Object tabParam) throws TabularDataException
    {
        if (fpath == null)
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "Source file path cannot be null.");
        }

        if (!Files.isRegularFile(fpath))
        {
            throw new TabularDataException(ErrorCode.FILE_NOT_FOUND, "File [" + fpath + "] does not exist.");
        }

        if (!fpath.toString().toLowerCase(Locale.ROOT).endsWith(".xlsx"))
        {
            throw new TabularDataException(ErrorCode.GENERAL_ERROR, "File [" + fpath + "] is not an XLSX file.");
        }

        if (tabParam != null && !(tabParam instanceof String) && !(tabParam instanceof Number))
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Worksheet [" + tabParam + "] is unknown or undefined.");
        }

        this.filePath = fpath;
        this.tabParam = tabParam;
        this.useStyleTable = useStyleTable;
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
     * Loads the configured worksheet into the reader.
     *
     * <p>
     * If a worksheet name or order index has been configured, that worksheet is selected.
     * Otherwise, the configured zero-based worksheet index (first tab) is selected by default.
     * </p>
     *
     * @throws TabularDataException
     *         if an error occurs while opening the workbook or reading the selected worksheet
     */
    @Override
    public void open() throws TabularDataException
    {
        if (tabParam instanceof String)
        {
            currentSheet = (String) tabParam;
        }

        else if (tabParam instanceof Integer)
        {
            int index = ((Integer) tabParam).intValue();
            List<String> sheets = fetchWorksheetList(filePath);

            if (index >= 0 && index < sheets.size())
            {
                currentSheet = sheets.get(index);
            }

            else
            {
                throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "Worksheet index [" + index + "] is out of bounds.");
            }
        }

        else
        {
            currentSheet = null;
        }
    }

    /**
     * Clears cached row data held by the current worksheet handler.
     *
     * <p>
     * Worksheet streams and workbook packages are closed by the worksheet-reading methods that open
     * them.
     * </p>
     */
    @Override
    public void close()
    {
        handler = null;
    }

    /**
     * Returns the number of rows cached for the active worksheet.
     *
     * @return cached row count, or {@code 0} if no worksheet handler is available
     */
    @Override
    public int getRowCount()
    {
        return (handler != null ? handler.getRowCount() : 0);
    }

    /**
     * Returns the maximum number of columns held by the active worksheet's parsed matrix.
     *
     * @return maximum matrix width, or {@code 0} if no worksheet handler is available
     */
    @Override
    public int getColumnCount()
    {
        return (handler != null ? handler.getMaximumMatrixWidth() : 0);
    }

    /**
     * Finds the first column in the specified row whose cell value equals the specified text,
     * ignoring case and leading or trailing whitespace.
     *
     * <p>
     * A {@code null} search value is treated as an empty string. Null cell values are skipped,
     * while empty-string cell values can match an empty search value.
     * </p>
     *
     * @param entry
     *        text to match or {@code null} is treated as an empty string
     * @param rowIndex
     *        zero-based index of the row to search
     * @return zero-based column index of the first matching cell, or {@code -1} if no match is
     *         found or the row index is invalid
     */
    @Override
    public int findColumnIndexByContent(String entry, int rowIndex)
    {
        // Streaming lookup for column index at a given row
        final int targetRow = rowIndex;
        final String searchEntry = entry != null ? entry.trim() : "";
        final int[] foundCol = {-1};

        try
        {
            Consumer<String[]> finderConsumer = new Consumer<String[]>()
            {
                private int currentRow = 0;

                @Override
                public void accept(String[] row)
                {
                    if (currentRow == targetRow)
                    {
                        for (int j = 0; j < row.length; j++)
                        {
                            if (row[j] != null && row[j].trim().equalsIgnoreCase(searchEntry))
                            {
                                foundCol[0] = j;
                                break;
                            }
                        }
                    }

                    currentRow++;
                }
            };

            if (tabParam instanceof String)
            {
                readWorksheet((String) tabParam, finderConsumer);
            }

            else if (tabParam instanceof Integer)
            {
                readWorksheet(((Integer) tabParam).intValue(), finderConsumer);
            }

            else
            {
                readWorksheet(0, finderConsumer);
            }
        }

        catch (Exception exc)
        {
            return -1;
        }

        return foundCol[0];
    }

    /**
     * Finds the first column in the specified row whose cell value matches the specified regular
     * expression, ignoring case.
     *
     * <p>
     * The cell value is trimmed before matching, and the expression must match the entire value. If
     * the expression is invalid, it is treated as literal text as a fallback value.
     * </p>
     *
     * @param regex
     *        regular expression to match, or literal text if the expression is invalid
     * @param rowIndex
     *        zero-based index of the row to search
     * @return zero-based column index of the first matching cell, or {@code -1} if no match is
     *         found, {@code regex} is {@code null}, or {@code rowIndex} is invalid
     */
    @Override
    public int findColumnIndexRegex(String regex, int rowIndex)
    {
        if (regex == null || rowIndex < 0)
        {
            return -1;
        }

        Pattern pattern;

        try
        {
            pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
        }

        catch (PatternSyntaxException exc)
        {
            pattern = Pattern.compile(Pattern.quote(regex), Pattern.CASE_INSENSITIVE);
        }

        final int targetRow = rowIndex;
        final Pattern targetPattern = pattern;
        final int[] foundCol = {-1};

        try
        {
            Consumer<String[]> finderConsumer = new Consumer<String[]>()
            {
                private int currentRow = 0;

                @Override
                public void accept(String[] row)
                {
                    if (currentRow == targetRow)
                    {
                        for (int j = 0; j < row.length; j++)
                        {
                            Matcher matcher = targetPattern.matcher(row[j] != null ? row[j].trim() : "");

                            if (matcher.matches())
                            {
                                foundCol[0] = j;
                                break;
                            }
                        }
                    }

                    currentRow++;
                }
            };

            if (tabParam instanceof String)
            {
                readWorksheet((String) tabParam, finderConsumer);
            }

            else if (tabParam instanceof Integer)
            {
                readWorksheet(((Integer) tabParam).intValue(), finderConsumer);
            }

            else
            {
                readWorksheet(0, finderConsumer);
            }
        }

        catch (Exception exc)
        {
            return -1;
        }

        return foundCol[0];
    }

    /**
     * Returns the name of the most recently selected worksheet.
     *
     * @return active worksheet name, or {@code null} if no worksheet has been selected successfully
     */
    @Override
    public String getWorksheetName()
    {
        return currentSheet;
    }

    /**
     * Returns the number of worksheets in the workbook.
     *
     * <p>
     * It may return {@code 0} if the worksheet list cannot be obtained. Therefore, a return value
     * of {@code 0} does not necessarily mean the workbook contains no worksheets.
     * </p>
     *
     * @return worksheet count, or {@code 0} if the worksheet list cannot be read
     */
    @Override
    public int getSheetCount()
    {
        try
        {
            return fetchWorksheetList(filePath).size();
        }

        catch (TabularDataException exc)
        {
            return 0;
        }
    }

    @Override
    public Iterator<String[]> iterator()
    {
        throw new UnsupportedOperationException("EventExcelReader operates in consumer-only streaming mode and does not support iteration.");
    }

    /**
     * Required by ExcelReadable interface (defaults to a no-op consumer if none provided).
     */
    @Override
    public void readWorksheet(String tabName) throws TabularDataException
    {
        readWorksheet(tabName, null);
    }

    /**
     * Required by ExcelReadable interface (defaults to a no-op consumer if none provided).
     */
    @Override
    public void readWorksheet(int tabIndex) throws TabularDataException
    {
        readWorksheet(tabIndex, null);
    }

    /**
     * - Reads the currently configured
     * worksheet, streaming each row to the provided consumer.
     *
     * -
     * 
     * @throws TabularDataException
     *         - if no
     *         worksheet has been specified or configured
     */
    public void readWorksheet(Consumer<String[]> rowConsumer) throws TabularDataException
    {
        if (currentSheet != null && !currentSheet.isEmpty())
        {
            readWorksheet(currentSheet, rowConsumer);
        }

        else if (tabParam instanceof String)
        {
            readWorksheet((String) tabParam, rowConsumer);
        }

        else if (tabParam instanceof Integer)
        {
            readWorksheet(((Integer) tabParam).intValue(), rowConsumer);
        }

        else
        {
            throw new TabularDataException(ErrorCode.SHEET_UNKNOWN, "No worksheet has been specified for this reader instance.");
        }
    }

    /**
     * Reads the specified worksheet and makes its rows available to this reader.
     *
     * <p>
     * Worksheet names are matched without regard to case. If {@code tabName} is {@code null} or
     * empty, the first worksheet is selected.
     * </p>
     *
     * @param tabName
     *        name of the worksheet to read, or {@code null} or an empty string to select the first
     *        worksheet
     *
     * @throws TabularDataException
     *         if an I/O, workbook-processing, or XML-parsing error occurs
     */
    private void readWorksheet(String tabName, Consumer<String[]> rowConsumer) throws TabularDataException
    {
        if (tabName == null || tabName.isEmpty())
        {
            readWorksheet(0, rowConsumer);
            return;
        }

        final Consumer<String[]> activeConsumer = (rowConsumer != null ? rowConsumer : new Consumer<String[]>()
        {
            @Override
            public void accept(String[] row)
            {
                // No-op
            }
        });

        try (OPCPackage pkg = OPCPackage.open(filePath.toFile(), PackageAccess.READ))
        {
            XSSFReader xssfReader = new XSSFReader(pkg);
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(pkg);

            handler = (useStyleTable ? new ExcelHandler(sst, xssfReader.getStylesTable(), activeConsumer) : new ExcelHandler(sst, activeConsumer));

            while (iter.hasNext())
            {
                try (InputStream stream = iter.next())
                {
                    if (tabName.equalsIgnoreCase(iter.getSheetName()))
                    {
                        currentSheet = iter.getSheetName();
                        parseSheet(stream);
                        break;
                    }
                }
            }
        }
        catch (IOException | OpenXML4JException | SAXException exc)
        {
            throw new TabularDataException(ErrorCode.STREAM_READ_FAILED, "Error streaming sheet [" + tabName + "] in [" + filePath + "]", exc);
        }
    }

    /**
     * Reads the worksheet at the specified zero-based index and makes its rows available for this
     * reader.
     *
     * @param tabIndex
     *        zero-based index of the worksheet to read
     *
     * @throws TabularDataException
     *         if an I/O, workbook-processing, or XML-parsing error occurs
     */
    private void readWorksheet(int tabIndex, Consumer<String[]> rowConsumer) throws TabularDataException
    {
        final Consumer<String[]> activeConsumer = (rowConsumer != null) ? rowConsumer : new Consumer<String[]>()
        {
            @Override
            public void accept(String[] row)
            {
                // No-op
            }
        };

        try (OPCPackage pkg = OPCPackage.open(filePath.toFile(), PackageAccess.READ))
        {
            int currentIndex = 0;
            XSSFReader xssfReader = new XSSFReader(pkg);
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(pkg);

            handler = (useStyleTable ? new ExcelHandler(sst, xssfReader.getStylesTable(), activeConsumer) : new ExcelHandler(sst, activeConsumer));

            while (iter.hasNext())
            {
                try (InputStream stream = iter.next())
                {
                    if (currentIndex == tabIndex)
                    {
                        currentSheet = iter.getSheetName();
                        parseSheet(stream);
                        break;
                    }
                }

                currentIndex++;
            }
        }

        catch (IOException | OpenXML4JException | SAXException exc)
        {
            throw new TabularDataException(ErrorCode.STREAM_READ_FAILED, "Error streaming sheet index [" + tabIndex + "] in [" + filePath + "]", exc);
        }
    }

    /**
     * Fetches a list of current available worksheet names defined in the workbook.
     *
     * <p>
     * It uses a StAX event reader to inspect the workbook metadata without building a DOM tree of
     * the XML document. Worksheet names are returned in workbook order.
     * </p>
     *
     * @param fpath
     *        to the XLSX workbook
     * @return worksheet names in workbook order
     *
     * @throws TabularDataException
     *         if the workbook cannot be opened or its worksheet metadata cannot be read
     */
    public static List<String> fetchWorksheetList(Path fpath) throws TabularDataException
    {
        List<String> sheetList = new ArrayList<>();

        try (OPCPackage opcpackage = OPCPackage.open(fpath.toFile(), PackageAccess.READ))
        {
            PackagePart workbookpart = opcpackage.getPartsByContentType(XSSFRelation.WORKBOOK.getContentType()).get(0);

            XMLInputFactory factory = XMLInputFactory.newInstance();
            factory.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, false);
            factory.setProperty(XMLInputFactory.SUPPORT_DTD, false);

            XMLEventReader reader = factory.createXMLEventReader(workbookpart.getInputStream());

            try
            {
                while (reader.hasNext())
                {
                    XMLEvent event = reader.nextEvent();

                    if (event.isStartElement())
                    {
                        StartElement startElement = event.asStartElement();
                        QName startElementName = startElement.getName();

                        if ("sheet".equalsIgnoreCase(startElementName.getLocalPart()))
                        {
                            Attribute attribute = startElement.getAttributeByName(new QName("name"));

                            if (attribute != null)
                            {
                                sheetList.add(attribute.getValue());
                            }
                        }
                    }
                }
            }

            finally
            {
                reader.close();
            }
        }

        catch (XMLStreamException | InvalidFormatException | IOException exc)
        {
            throw new TabularDataException(ErrorCode.STREAM_READ_FAILED, "Failed to load worksheet list from [" + fpath + "]", exc);
        }

        return sheetList;
    }

    /**
     * Parses worksheet XML using a SAX parser and passes the parsed events to the current worksheet
     * handler.
     *
     * @param stream
     *        input stream containing worksheet XML
     * @throws SAXException
     *         if SAX parsing fails
     *
     * @throws IOException
     *         if an I/O error occurs while parsing the stream
     * @throws IllegalStateException
     *         if the SAX parser cannot be configured
     */
    private void parseSheet(InputStream stream) throws SAXException, IOException
    {
        InputSource sheetSource = new InputSource(stream);

        try
        {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        }

        catch (ParserConfigurationException exc)
        {
            throw new IllegalStateException("SAX parser configuration error: " + exc.getMessage(), exc);
        }
    }

    public static void main(String[] args) throws TabularDataException
    {
        Path file = Paths.get("src/io/test", "Praveen_batch_list.xlsx");

        // Instantiate without a default sheet to test late binding / exception safety
        try (EventExcelReader reader = new EventExcelReader(file, true))
        {
            reader.open();

            // Stream rows from a specifically named worksheet
            reader.readWorksheet("qantas.com", new Consumer<String[]>()
            {
                @Override
                public void accept(String[] row)
                {
                    System.out.println(Arrays.toString(row));
                }
            });

            System.out.printf("Active Sheet: %s%n", reader.getWorksheetName());

            // Switch and stream another worksheet ("Sheet1") by name
            reader.readWorksheet("Sheet1", new Consumer<String[]>()
            {
                @Override
                public void accept(String[] row)
                {
                    System.out.println(Arrays.toString(row));
                }
            });
        }
        catch (TabularDataException exc)
        {
            System.err.printf("Engine error [%s]: %s%n", exc.getErrorCode(), exc.getMessage());
            exc.printStackTrace();
        }
    }
}