package io.metadata;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Optional;
import java.util.Properties;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.xmlbeans.XmlException;

/**
 * Reads the core document properties stored within an OOXML workbook or spreadsheet.
 *
 * <p>
 * The workbook is opened using Apache POI's {@link OPCPackage} API rather than loading the complete
 * workbook into memory. This approach minimises memory usage when only document metadata is
 * required, particularly when processing large workbooks.
 * </p>
 *
 * <p>
 * The class provides access to common core properties such as title, subject, creator, description,
 * keywords, creation date, modification date, and revision number. Missing string properties and
 * formatted date values are returned as empty strings, whilst missing date values are returned as
 * {@code null}.
 * </p>
 *
 * <b>Change log:</b>
 * <ul>
 * <li>22 February 2018 - Initial creation.</li>
 * <li>6 October 2026 - Refactored date formatting to use DateTimeFormatter and updated
 * documentation to British spelling.</li>
 * </ul>
 *
 * @author Trevor Maggs
 * @version 0.2
 * @since 22 February 2018
 */
public class ExcelCoreProperties
{
    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd-MM-yyyy HH:mm");
    private final POIXMLProperties.CoreProperties coreProp;

    /**
     * Creates an {@code ExcelCoreProperties} instance for the specified workbook and reads its core
     * document properties.
     *
     * @param fpath
     *        path to the workbook or spreadsheet file
     *
     * @throws IOException
     *         if the file does not exist or its core properties cannot be read
     */
    public ExcelCoreProperties(Path fpath) throws IOException
    {
        if (!Files.isRegularFile(fpath))
        {
            throw new FileNotFoundException("File [" + fpath + "] does not exist or is not a regular file.");
        }

        try (OPCPackage opcpackage = OPCPackage.open(fpath.toFile()))
        {
            POIXMLProperties metadata = new POIXMLProperties(opcpackage);
            coreProp = metadata.getCoreProperties();
        }

        catch (InvalidFormatException exc)
        {
            throw new IOException("File [" + fpath + "] does not exist or has parsing errors.", exc);
        }

        catch (OpenXML4JException | XmlException exc)
        {
            throw new IOException("Problem reading core properties in file [" + fpath + "].", exc);
        }
    }

    /**
     * Returns the value contained in the specified {@link Optional} string.
     *
     * <p>
     * An empty optional is represented by an empty string.
     * </p>
     *
     * @param entry
     *        optional containing the string property value
     * @return the property value, or an empty string if the value is unavailable
     */
    private String getStringValue(Optional<String> entry)
    {
        return entry.orElse("");
    }

    /**
     * Formats the date contained in the specified {@link Optional}.
     *
     * <p>
     * An empty optional is represented by an empty string.
     * </p>
     *
     * @param entry
     *        optional containing the date property value
     * @return the formatted date, or an empty string if the value is unavailable
     */
    private String formatDateString(Optional<Date> entry)
    {
        if (entry.isPresent())
        {
            return DATE_FORMATTER.format(entry.get().toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
        }

        return "";
    }

    /**
     * Returns the date contained in the specified {@link Optional}.
     *
     * <p>
     * An empty optional results in a {@code null} return value.
     * </p>
     *
     * @param entry
     *        optional containing the date property value
     * @return the date as an {@link Instant}, or {@code null} if the value is unavailable
     */
    private Instant getDate(Optional<Date> entry)
    {
        return (entry.isPresent() ? entry.get().toInstant() : null);
    }

    /**
     * Returns the category assigned to the workbook.
     *
     * @return the category, or an empty string if unavailable
     */
    public String getCategory()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getCategoryProperty());
    }

    /**
     * Returns the content status of the workbook.
     *
     * @return the content status, or an empty string if unavailable
     */
    public String getContentStatus()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getContentStatusProperty());
    }

    /**
     * Returns the content type of the workbook.
     *
     * @return the content type, or an empty string if unavailable
     */
    public String getContentType()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getContentTypeProperty());
    }

    /**
     * Returns the date on which the workbook was created.
     *
     * @return the formatted creation date, or an empty string if unavailable
     */
    public String getDateCreated()
    {
        return formatDateString(coreProp.getUnderlyingProperties().getCreatedProperty());
    }

    /**
     * Returns the date on which the workbook was created.
     *
     * @return the creation date, or {@code null} if unavailable
     */
    public Instant getDateCreatedValue()
    {
        return getDate(coreProp.getUnderlyingProperties().getCreatedProperty());
    }

    /**
     * Returns the name of the user who created the workbook.
     *
     * @return the creator's name, or an empty string if unavailable
     */
    public String getCreator()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getCreatorProperty());
    }

    /**
     * Returns the description associated with the workbook.
     *
     * @return the description, or an empty string if unavailable
     */
    public String getDescription()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getDescriptionProperty());
    }

    /**
     * Returns the identifier assigned to the workbook.
     *
     * @return the identifier, or an empty string if unavailable
     */
    public String getIdentifier()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getIdentifierProperty());
    }

    /**
     * Returns the keywords associated with the workbook.
     *
     * @return the keywords, or an empty string if unavailable
     */
    public String getKeywords()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getKeywordsProperty());
    }

    /**
     * Returns the name of the user who last modified the workbook.
     *
     * @return the last modifying user's name, or an empty string if unavailable
     */
    public String getLastModifiedByUser()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getLastModifiedByProperty());
    }

    /**
     * Returns the date on which the workbook was last printed.
     *
     * @return the formatted last-printed date, or an empty string if unavailable
     */
    public String getLastPrinted()
    {
        return formatDateString(coreProp.getUnderlyingProperties().getLastPrintedProperty());
    }

    /**
     * Returns the date on which the workbook was last modified.
     *
     * @return the formatted modification date, or an empty string if unavailable
     */
    public String getDateModified()
    {
        return formatDateString(coreProp.getUnderlyingProperties().getModifiedProperty());
    }

    /**
     * Returns the revision number of the workbook.
     *
     * @return the revision number, or an empty string if unavailable
     */
    public String getRevision()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getRevisionProperty());
    }

    /**
     * Returns the subject associated with the workbook.
     *
     * @return the subject, or an empty string if unavailable
     */
    public String getSubject()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getSubjectProperty());
    }

    /**
     * Returns the title of the workbook.
     *
     * @return the title, or an empty string if unavailable
     */
    public String getTitle()
    {
        return getStringValue(coreProp.getUnderlyingProperties().getTitleProperty());
    }

    /**
     * Returns the workbook's core document properties as name/value pairs.
     *
     * <p>
     * The returned {@link Properties} object contains the available core properties, including
     * category, content status, content type, creation date, creator, description, identifier,
     * keywords, last modifying user, last printed date, modification date, revision, subject, and
     * title. Missing values are represented by empty strings in the returned {@link Properties}
     * object.
     * </p>
     *
     * @return a new {@link Properties} object containing the workbook's core document properties
     */
    public Properties getProperties()
    {
        Properties prop = new Properties();

        prop.setProperty("category", getCategory());
        prop.setProperty("contentStatus", getContentStatus());
        prop.setProperty("contentType", getContentType());
        prop.setProperty("dateCreated", getDateCreated());
        prop.setProperty("creator", getCreator());
        prop.setProperty("description", getDescription());
        prop.setProperty("identifier", getIdentifier());
        prop.setProperty("keywords", getKeywords());
        prop.setProperty("lastModifiedByUser", getLastModifiedByUser());
        prop.setProperty("dateLastPrinted", getLastPrinted());
        prop.setProperty("dateModified", getDateModified());
        prop.setProperty("revision", getRevision());
        prop.setProperty("subject", getSubject());
        prop.setProperty("title", getTitle());

        return prop;
    }
}