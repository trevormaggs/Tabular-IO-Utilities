package io.metadata;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Properties;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

/**
 * Reads and writes additional custom document properties stored within an OOXML workbook or
 * spreadsheet.
 *
 * <p>
 * The workbook is accessed using Apache POI's {@link OPCPackage} API rather than loading it into a
 * full XSSFWorkbook object model. This avoids the overhead of constructing the complete workbook
 * model when only document metadata is required.
 * </p>
 *
 * <b>Change log:</b>
 * <ul>
 * <li>16 November 2021 - Initial creation.</li>
 * <li>7 October 2026 - Refactored package location to io.metadata and improved property value
 * extraction.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.3
 * @since 16 November 2021
 */
public class ExcelCustomProperties
{
    private final Path workfile;

    /**
     * Creates a {@code ExcelCustomProperties} instance for the specified workbook.
     *
     * @param pfile
     *        path to the workbook or spreadsheet file
     *
     * @throws IOException
     *         if the file does not exist
     */
    public ExcelCustomProperties(Path pfile) throws IOException
    {
        if (Files.notExists(pfile))
        {
            throw new FileNotFoundException("File [" + pfile + "] does not exist.");
        }

        workfile = pfile;
    }

    /**
     * Adds custom properties to the workbook.
     *
     * @param prop
     *        a {@link Properties} object containing custom name/value pairs
     *
     * @throws IOException
     *         if the file cannot be accessed or custom properties cannot be written
     */
    public void setProperties(Properties prop) throws IOException
    {
        if (prop == null || prop.isEmpty())
        {
            return;
        }

        try (OPCPackage opcPackage = OPCPackage.open(workfile.toFile()))
        {
            POIXMLProperties metadata = new POIXMLProperties(opcPackage);
            POIXMLProperties.CustomProperties custProp = metadata.getCustomProperties();

            for (String key : prop.stringPropertyNames())
            {
                custProp.addProperty(key, prop.getProperty(key));
            }

            metadata.commit();
        }

        catch (InvalidFormatException exc)
        {
            throw new IOException("File [" + workfile + "] does not exist or has parsing errors.", exc);
        }

        catch (OpenXML4JException | XmlException exc)
        {
            throw new IOException("Problem writing custom properties in file [" + workfile + "].", exc);
        }
    }

    /**
     * Retrieves all custom document properties stored within the workbook.
     *
     * @return a new {@link Properties} object containing custom property name/value pairs
     *
     * @throws IOException
     *         if the file cannot be accessed or custom properties cannot be read
     */
    public Properties getProperties() throws IOException
    {
        Properties prop = new Properties();

        try (OPCPackage opcPackage = OPCPackage.open(workfile.toFile()))
        {
            POIXMLProperties metadata = new POIXMLProperties(opcPackage);
            POIXMLProperties.CustomProperties custProp = metadata.getCustomProperties();
            List<CTProperty> propList = custProp.getUnderlyingProperties().getPropertyList();

            for (CTProperty pItem : propList)
            {
                String propertyName = pItem.getName();
                String propertyValue = extractPropertyValue(pItem);

                if (propertyName != null)
                {
                    prop.setProperty(propertyName, propertyValue);
                }
            }

            opcPackage.revert();
        }

        catch (InvalidFormatException exc)
        {
            throw new IOException("File [" + workfile + "] does not exist or has parsing errors.", exc);
        }

        catch (OpenXML4JException | XmlException exc)
        {
            throw new IOException("Problem reading custom properties in file [" + workfile + "].", exc);
        }

        return prop;
    }

    /**
     * Extracts a string representation from a supported custom property XML element.
     * 
     * @param pItem
     *        the underlying XML property element
     * @return the property value as a string, or an empty string if the property type is
     *         unsupported
     */
    private String extractPropertyValue(CTProperty pItem)
    {
        if (pItem.isSetLpwstr())
        {
            return pItem.getLpwstr();
        }

        if (pItem.isSetI4())
        {
            return Integer.toString(pItem.getI4());
        }

        if (pItem.isSetR8())
        {
            return Double.toString(pItem.getR8());
        }

        if (pItem.isSetBool())
        {
            return Boolean.toString(pItem.getBool());
        }

        if (pItem.isSetFiletime())
        {
            return pItem.getFiletime().toString();
        }

        return "";
    }
}