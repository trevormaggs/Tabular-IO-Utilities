package io;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public final class SpreadsheetFileDetector
{
    private static final String ODS_MIME_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

    private static final byte[] OLE_SIGNATURE = {
            (byte) 0xD0, (byte) 0xCF, (byte) 0x11, (byte) 0xE0,
            (byte) 0xA1, (byte) 0xB1, (byte) 0x1A, (byte) 0xE1
    };

    private static final byte[] ZIP_SIGNATURE = {
            (byte) 0x50, (byte) 0x4B, (byte) 0x03, (byte) 0x04
    };

    public enum SpreadsheetFormat
    {
        XLS, XLSX, ODS, UNKNOWN
    }

    public enum BiffVersion
    {
        BIFF2, BIFF3, BIFF4, BIFF5, BIFF8, UNKNOWN
    }

    public static final class DetectionResult
    {
        private final SpreadsheetFormat format;
        private final BiffVersion biffVersion;
        private final boolean poiValidated;

        private DetectionResult(SpreadsheetFormat format, BiffVersion biffVersion, boolean poiValidated)
        {
            this.format = format;
            this.biffVersion = biffVersion;
            this.poiValidated = poiValidated;
        }

        public SpreadsheetFormat getFormat()
        {
            return format;
        }

        public BiffVersion getBiffVersion()
        {
            return biffVersion;
        }

        /**
         * Returns whether Apache POI successfully parsed the workbook. For ODS, this is always
         * false because standard POI workbook parsing does not support the OpenDocument Spreadsheet
         * format.
         */
        public boolean isPoiValidated()
        {
            return poiValidated;
        }
    }

    private SpreadsheetFileDetector()
    {
        // Utility class.
    }

    /**
     * Detects XLS, XLSX and ODS formats using file contents.
     *
     * @param path
     *        path to the spreadsheet
     * @return the detection result
     * 
     * @throws IOException
     *         if an I/O error occurs during initial detection
     */
    public static DetectionResult detect(Path path) throws IOException
    {

        if (path == null || !Files.isRegularFile(path) || !Files.isReadable(path))
        {
            return result(SpreadsheetFormat.UNKNOWN, BiffVersion.UNKNOWN, false);
        }

        // 1. Legacy Excel: OLE container containing a BIFF stream.
        if (hasSignature(path, OLE_SIGNATURE))
        {
            BiffVersion version = detectBiffVersion(path);

            if (version != BiffVersion.UNKNOWN)
            {
                return result(SpreadsheetFormat.XLS, version, isPoiValidWorkbook(path));
            }
        }

        // 2. ZIP-based formats: XLSX or ODS.
        if (hasSignature(path, ZIP_SIGNATURE))
        {
            if (isOdsFile(path))
            {
                // ODS is recognised from its package structure.
                // Apache POI's standard WorkbookFactory does not
                // validate ODS workbook contents.
                return result(SpreadsheetFormat.ODS, BiffVersion.UNKNOWN, false);
            }

            if (isXlsxPackage(path))
            {
                return result(SpreadsheetFormat.XLSX, BiffVersion.UNKNOWN, isPoiValidWorkbook(path));
            }
        }

        return result(SpreadsheetFormat.UNKNOWN, BiffVersion.UNKNOWN, false);
    }

    private static DetectionResult result(SpreadsheetFormat format, BiffVersion version, boolean poiValidated)
    {
        return new DetectionResult(format, version, poiValidated);
    }

    /**
     * Compares the file's initial bytes with a signature.
     */
    private static boolean hasSignature(Path path, byte[] signature) throws IOException
    {
        byte[] header = new byte[signature.length];

        try (InputStream in = Files.newInputStream(path))
        {
            int offset = 0;

            while (offset < header.length)
            {
                int count = in.read(header, offset, header.length - offset);

                if (count < 0)
                {
                    return false;
                }

                offset += count;
            }
        }

        for (int i = 0; i < signature.length; i++)
        {
            if (header[i] != signature[i])
            {
                return false;
            }
        }

        return true;
    }

    /**
     * Finds the BIFF workbook stream and determines its BIFF version.
     */
    private static BiffVersion detectBiffVersion(Path path) throws IOException
    {
        try (POIFSFileSystem fs = new POIFSFileSystem(path.toFile(), true))
        {
            String streamName;

            if (fs.getRoot().hasEntry("Workbook"))
            {
                streamName = "Workbook";
            }

            else if (fs.getRoot().hasEntry("Book"))
            {
                streamName = "Book";
            }

            else
            {
                return BiffVersion.UNKNOWN;
            }

            try (InputStream in = fs.getRoot().createDocumentInputStream(streamName))
            {

                int idLow = in.read();
                int idHigh = in.read();
                int lenLow = in.read();
                int lenHigh = in.read();

                if (idLow < 0 || idHigh < 0 || lenLow < 0 || lenHigh < 0)
                {
                    return BiffVersion.UNKNOWN;
                }

                int recordId = idLow | (idHigh << 8);
                int recordLength = lenLow | (lenHigh << 8);

                if (recordLength < 4)
                {
                    return BiffVersion.UNKNOWN;
                }

                switch (recordId)
                {
                    case 0x0009:
                        return BiffVersion.BIFF2;

                    case 0x0209:
                        return BiffVersion.BIFF3;

                    case 0x0409:
                        return BiffVersion.BIFF4;

                    case 0x0809:
                        return detectBiff5Or8(in);

                    default:
                        return BiffVersion.UNKNOWN;
                }
            }
        }
    }

    /**
     * Reads the BOF version field following the four-byte record header.
     */
    private static BiffVersion detectBiff5Or8(InputStream in) throws IOException
    {
        int versionLow = in.read();
        int versionHigh = in.read();

        if (versionLow < 0 || versionHigh < 0)
        {
            return BiffVersion.UNKNOWN;
        }

        int version = versionLow | (versionHigh << 8);

        switch (version)
        {
            case 0x0500:
                return BiffVersion.BIFF5;

            case 0x0600:
                return BiffVersion.BIFF8;

            default:
                return BiffVersion.UNKNOWN;
        }
    }

    /**
     * Recognises ODS by its MIME type and expected content entry.
     */
    private static boolean isOdsFile(Path path)
    {
        try (ZipFile zip = new ZipFile(path.toFile()))
        {
            ZipEntry mimeEntry = zip.getEntry("mimetype");
            ZipEntry contentEntry = zip.getEntry("content.xml");

            if (mimeEntry == null || mimeEntry.isDirectory() || contentEntry == null || contentEntry.isDirectory())
            {
                return false;
            }

            if (mimeEntry.getSize() > 256)
            {
                return false;
            }

            // A conforming ODS package stores mimetype first and uncompressed.
            Enumeration<? extends ZipEntry> entries = zip.entries();

            if (!entries.hasMoreElements())
            {
                return false;
            }

            ZipEntry firstEntry = entries.nextElement();

            if (!"mimetype".equals(firstEntry.getName()) || firstEntry.getMethod() != ZipEntry.STORED)
            {
                return false;
            }

            try (InputStream in = zip.getInputStream(mimeEntry); ByteArrayOutputStream out = new ByteArrayOutputStream())
            {
                int count;
                byte[] buffer = new byte[64];

                while ((count = in.read(buffer)) != -1)
                {
                    if (out.size() + count > 256)
                    {
                        return false;
                    }

                    out.write(buffer, 0, count);
                }

                String mimeType = new String(out.toByteArray(), StandardCharsets.US_ASCII);

                return ODS_MIME_TYPE.equals(mimeType);
            }

        }

        catch (IOException | RuntimeException e)
        {
            return false;
        }
    }

    /**
     * Checks the basic package entries expected in an XLSX file.
     */
    private static boolean isXlsxPackage(Path path)
    {

        try (ZipFile zip = new ZipFile(path.toFile()))
        {
            ZipEntry contentTypes = zip.getEntry("[Content_Types].xml");
            ZipEntry workbookXml = zip.getEntry("xl/workbook.xml");

            return contentTypes != null && workbookXml != null && !contentTypes.isDirectory() && !workbookXml.isDirectory();
        }
        
        catch (IOException | RuntimeException e)
        {
            return false;
        }
    }

    /**
     * Uses Apache POI to attempt XLS/XLSX workbook parsing.
     */
    private static boolean isPoiValidWorkbook(Path path)
    {

        try (Workbook workbook = WorkbookFactory.create(path.toFile()))
        {
            return workbook.getNumberOfSheets() > 0;
        }
        
        catch (IOException | RuntimeException e)
        {
            return false;
        }
    }

    /**
     * Command-line example.
     */
    public static void main(String[] args)
    {
        if (args.length != 1)
        {
            System.out.println("Usage: java SpreadsheetFileDetector <file>");
            return;
        }

        try
        {
            DetectionResult result = detect(Paths.get(args[0]));

            System.out.println("Format: " + result.getFormat());

            if (result.getFormat() == SpreadsheetFormat.XLS)
            {
                System.out.println("BIFF version: " + result.getBiffVersion());
            }

            if (result.getFormat() == SpreadsheetFormat.ODS)
            {
                System.out.println("Apache POI validation: not applicable");
            }
            
            else
            {
                System.out.println("Apache POI validation: " + result.isPoiValidated());
            }
        }
        
        catch (IOException e)
        {
            System.err.println("Unable to read file: " + e.getMessage());
        }
    }
}