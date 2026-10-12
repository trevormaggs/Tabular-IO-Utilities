package io.test;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;
import javax.xml.namespace.QName;
import javax.xml.stream.FactoryConfigurationError;
import javax.xml.stream.XMLEventFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLEventWriter;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

public class ReplaceExcelSST
{
    private final OPCPackage opcpackage;
    private final XMLEventFactory eventFactory;
    private final SharedStringsTable sharedStringsTable;
    private XMLEventReader readstream;
    private XMLEventWriter writestream;
    private String currentData;
    private StartElement cellFound;

    /**
     * Main constructor to open the XML file for processing.
     * 
     * @param file
     *        relative or absolute path to the file
     * 
     * @throws IOException
     *         if there is any problem
     */
    public ReplaceExcelSST(String file) throws IOException
    {
        sharedStringsTable = new SharedStringsTable();
        eventFactory = XMLEventFactory.newInstance();

        if (!file.toLowerCase().endsWith(".xlsx"))
        {
            throw new IllegalArgumentException("File [" + file + "] is not an XML-formatted type. Please provide an XLSX excel file.");
        }

        try
        {
            opcpackage = OPCPackage.open(file);
        }

        catch (InvalidFormatException exc)
        {
            throw new IOException("File [" + file + "] does not exist or has parsing errors.", exc);
        }
    }

    /**
     * Scans every cell iteratively located in the first worksheet and replaces inline strings with
     * shared strings.
     * 
     * <p>
     * <b>Important:</b> be mindful this method focuses on processing the first worksheet of the
     * spreadsheet only.
     * </p>
     * 
     * @throws IOException
     *         if there is any problem
     */
    private void replaceXMLEvent() throws IOException
    {
        PackagePart sheetpart = opcpackage.getPartsByName(Pattern.compile("/xl/worksheets/sheet1.xml")).get(0);
        PackagePart sstpart = opcpackage.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType()).get(0);

        // Instantiate and set to handle special HTML characters
        XMLInputFactory factory = XMLInputFactory.newInstance();
        factory.setProperty(XMLInputFactory.IS_COALESCING, true);

        try
        {
            sharedStringsTable.readFrom(sstpart.getInputStream());
            readstream = factory.createXMLEventReader(sheetpart.getInputStream());
            writestream = XMLOutputFactory.newInstance().createXMLEventWriter(sheetpart.getOutputStream(), "utf-8");

            // loop over all XML in sheet1.xml
            while (readstream.hasNext())
            {
                XMLEvent event = readstream.nextEvent();

                if (event.isStartDocument())
                {
                    writestream.add(event);
                }

                else if (event.isStartElement())
                {
                    processStart(event);

                    if (cellFound == null)
                    {
                        writestream.add(event);
                    }

                    handleStartElement();
                }

                else if (event.isCharacters())
                {
                    Characters data = event.asCharacters();
                    currentData = data.getData();

                    if (cellFound == null)
                    {
                        writestream.add(event);
                    }

                    handleCharacters();
                }

                else if (event.isEndElement())
                {
                    if (cellFound == null)
                    {
                        writestream.add(event);
                    }

                    processEnd(event);
                    handleEndElement();
                }
            }

            writestream.flush();

            // write the SharedStringsTable
            OutputStream out = sstpart.getOutputStream();

            sharedStringsTable.writeTo(out);
            out.close();
            opcpackage.close();
        }

        catch (XMLStreamException | FactoryConfigurationError exc)
        {
            throw new IOException(exc.getLocalizedMessage(), exc);
        }
    }

    /**
     * Receives notification of the start of an element for processing. If the
     * event identifies the current cell containing the inlineStr value, then it
     * will be marked for replacement.
     *
     * @param event
     *        the XMLEvent event pointing to the current element being
     *        seen
     * @throws XMLStreamException
     *         if there is any streaming problem
     */
    private void processStart(XMLEvent event) throws XMLStreamException
    {
        StartElement element = event.asStartElement();
        Attribute attr = element.getAttributeByName(new QName("t"));

        if (attr != null && "c".equals(element.getName().getLocalPart()))
        {
            /*
             * Example
             * <c r="AP2" s="2" t="inlineStr">
             * <is>
             * <t>Software upgrade</t>
             * </is>
             * </c>
             */

            String cellType = attr.getValue();

            if ("inlineStr".equals(cellType))
            {
                currentData = "";
                cellFound = element;
            }
        }
    }

    /**
     * Receives notification of the end of an element for processing. If the
     * inlineStr value is identified in the current iteration, it will be
     * replaced with the shared string.
     *
     * @param event
     *        the XMLEvent event pointing to the current element being
     *        seen
     * @throws XMLStreamException
     *         if there is any streaming problem
     */
    private void processEnd(XMLEvent event) throws XMLStreamException
    {
        EndElement element = event.asEndElement();

        if (cellFound != null && "c".equals(element.getName().getLocalPart()))
        {
            // create the new cell element having the shared string
            List<Attribute> attrList;
            Attribute t = eventFactory.createAttribute("t", "s");
            Attribute r = cellFound.getAttributeByName(new QName("r"));
            Attribute s = cellFound.getAttributeByName(new QName("s"));

            if (r == null && s == null) attrList = Arrays.asList(new Attribute[]{t});
            else if (r == null) attrList = Arrays.asList(new Attribute[]{s, t});
            else if (s == null) attrList = Arrays.asList(new Attribute[]{r, t});
            else attrList = Arrays.asList(new Attribute[]{r, s, t});

            CTRst ctstr = CTRst.Factory.newInstance();
            ctstr.setT(currentData);

            StartElement newCellStart = eventFactory.createStartElement(new QName("c"), attrList.iterator(), null);
            writestream.add(newCellStart);

            StartElement newCellValue = eventFactory.createStartElement(new QName("v"), null, null);
            writestream.add(newCellValue);

            // int sRef = sharedStringsTable.addEntry(ctstr);
            // int sRef = sharedStringsTable.addSharedStringItem((RichTextString) ctstr);

            int sRef = sharedStringsTable.addSharedStringItem(new XSSFRichTextString(ctstr));

            Characters value = eventFactory.createCharacters(Integer.toString(sRef));
            writestream.add(value);

            EndElement newCellValueEnd = eventFactory.createEndElement(new QName("v"), null);
            writestream.add(newCellValueEnd);

            EndElement newCellEnd = eventFactory.createEndElement(new QName("c"), null);
            writestream.add(newCellEnd);

            cellFound = null;
        }
    }

    /**
     * Handles the current StartElement event. The access modifier for this
     * method is protected therefore you can override it in the subclass if
     * needed.
     */
    protected void handleStartElement()
    {
    }

    /**
     * Handles the current Characters event. The access modifier for this method
     * is protected therefore you can override it in the subclass if needed.
     */
    protected void handleCharacters()
    {
    }

    /**
     * Handles the current EndElement event. The access modifier for this method
     * is protected therefore you can override it in the subclass if needed.
     */
    protected void handleEndElement()
    {
    }

    /**
     * Executes the replacement process.
     * 
     * @throws IOException
     *         if there is any problem
     */
    public void replace() throws IOException
    {
        replaceXMLEvent();
    }

    public static void main(String[] args)
    {
        long startTime = System.nanoTime();

        ReplaceExcelSST sst;
        String file = args[0];

        try
        {
            sst = new ReplaceExcelSST(file);
            sst.replace();
        }

        catch (RuntimeException | IOException exc)
        {
            System.err.println(exc.getMessage());
        }

        System.out.format("Time: %.6f%n", ((double) (System.nanoTime() - startTime) / 1E9));
    }
}