package io.test;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;
import javax.xml.stream.XMLEventFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLEventWriter;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.XMLEvent;
import javax.xml.namespace.QName;
import java.io.File;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;
import java.util.Iterator;

class StaxReplaceInlineStrings
{
    public static void main(String[] args)
    {
        try
        {
            File file = new File("test.xlsx");
            OPCPackage opcpackage = OPCPackage.open(file);

            // if there are strings in the sheet data, we need the SharedStringsTable
            PackagePart sharedstringstablepart = opcpackage.getPartsByName(Pattern.compile("/xl/sharedStrings.xml")).get(0);
            SharedStringsTable sharedstringstable = new SharedStringsTable();

            sharedstringstable.readFrom(sharedstringstablepart.getInputStream());

            PackagePart sheetpart = opcpackage.getPartsByName(Pattern.compile("/xl/worksheets/sheet1.xml")).get(0);

            XMLEventReader reader = XMLInputFactory.newInstance().createXMLEventReader(sheetpart.getInputStream());
            XMLEventWriter writer = XMLOutputFactory.newInstance().createXMLEventWriter(sheetpart.getOutputStream());

            XMLEventFactory eventFactory = XMLEventFactory.newInstance();

            // loop over all XML in sheet1.xml
            while (reader.hasNext())
            {
                // marker whether cell having inline string was replaced by cell
                // having shared string
                boolean cellReplaced = false;

                XMLEvent event = (XMLEvent) reader.next();

                if (event.isStartElement())
                {
                    StartElement startElement = (StartElement) event;
                    QName startElementName = startElement.getName();

                    // start element of cell
                    if (startElementName.getLocalPart().equalsIgnoreCase("c"))
                    {
                        Attribute attribute;

                        // remember cell start
                        StartElement cellStart = startElement;

                        // get cell's attributes
                        Iterator<?> attributeIterator = cellStart.getAttributes();

                        while (attributeIterator.hasNext())
                        {
                            attribute = (Attribute) attributeIterator.next();

                            // cell has type attribute
                            if ("t".equals(attribute.getName().getLocalPart()))
                            {
                                String tvalue = attribute.getValue();

                                // cell type is inline string
                                if ("inlineStr".equals(tvalue))
                                {
                                    String inlineString = "";

                                    // read next start element - error if is not
                                    // a start element
                                    startElement = (StartElement) (XMLEvent) reader.next();
                                    startElementName = startElement.getName();

                                    // start element of inline string
                                    if (startElementName.getLocalPart().equalsIgnoreCase("is"))
                                    {
                                        // read next start element - error if is
                                        // not a start element
                                        startElement = (StartElement) (XMLEvent) reader.next();

                                        startElementName = startElement.getName();

                                        // start element of text
                                        if (startElementName.getLocalPart().equalsIgnoreCase("t"))
                                        {
                                            // read next characters element -
                                            // error if is not a characters
                                            // element
                                            Characters characters = (Characters) (XMLEvent) reader.next();

                                            // get text data
                                            inlineString = characters.getData();

                                            System.out.println(inlineString);
                                        }
                                    }

                                    // create shared string in shared strings
                                    // table
                                    CTRst ctstr = CTRst.Factory.newInstance();
                                    ctstr.setT(inlineString);

                                    // int sRef = sharedstringstable.addEntry(ctstr);
                                    int sRef = sharedstringstable.addSharedStringItem((RichTextString) ctstr);

                                    // we are replacing the cell element so skip
                                    // elements until end element of cell
                                    while (reader.hasNext())
                                    {
                                        event = (XMLEvent) reader.next();

                                        if (event.isEndElement())
                                        {
                                            EndElement endElement = (EndElement) event;
                                            QName endElementName = endElement.getName();

                                            // end element of cell
                                            if (endElementName.getLocalPart().equalsIgnoreCase("c"))
                                            {
                                                break;
                                            }
                                        }
                                    }

                                    // create the new cell element having the shared string
                                    Attribute r = cellStart.getAttributeByName(new QName("r"));
                                    Attribute s = cellStart.getAttributeByName(new QName("s"));
                                    Attribute t = eventFactory.createAttribute("t", "s");
                                    List<Attribute> attributeList = Arrays.asList(new Attribute[]{t});

                                    if (r != null && s != null)
                                    {
                                        attributeList = Arrays.asList(new Attribute[]{r, s, t});
                                    }

                                    else if (r != null)
                                    {
                                        attributeList = Arrays.asList(new Attribute[]{r, t});
                                    }

                                    else if (s != null)
                                    {
                                        attributeList = Arrays.asList(new Attribute[]{s, t});
                                    }

                                    System.out.println(attributeList);

                                    StartElement newCellStart = eventFactory.createStartElement(new QName("c"), attributeList.iterator(), null);
                                    writer.add(newCellStart);

                                    StartElement newCellValue = eventFactory.createStartElement(new QName("v"), null, null);
                                    writer.add(newCellValue);

                                    Characters value = eventFactory.createCharacters(Integer.toString(sRef));
                                    writer.add(value);

                                    EndElement newCellValueEnd = eventFactory.createEndElement(new QName("v"), null);
                                    writer.add(newCellValueEnd);

                                    EndElement newCellEnd = eventFactory.createEndElement(new QName("c"), null);
                                    writer.add(newCellEnd);

                                    // mark that cell was replaced
                                    cellReplaced = true;
                                    break;
                                }
                            }
                        }
                    }
                }

                if (!cellReplaced)
                {
                    // by default write each read event, except cell was replaced
                    writer.add(event);
                }
            }

            writer.flush();

            // write the SharedStringsTable
            OutputStream out = sharedstringstablepart.getOutputStream();

            sharedstringstable.writeTo(out);
            out.close();

            sharedstringstable.close();
            opcpackage.close();
        }

        catch (Exception exc)
        {
            exc.printStackTrace();
        }
    }
}