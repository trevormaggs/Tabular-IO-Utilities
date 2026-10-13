package io.test;

import java.io.BufferedInputStream;
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
    // See: https://docs.oasis-open.org/office/OpenDocument/v1.3/OpenDocument-v1.3-part3-schema.html
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
         * Returns whether Apache POI successfully opened the workbook and the workbook contains at
         * least one sheet. Returns {@code false} for ODS files because validation is not supported
         * for that format.
         */
        public boolean isPoiValidated()
        {
            return poiValidated;
        }
    }

    /**
     * Prevents direct instantiation.
     *
     * @throws UnsupportedOperationException
     *         to indicate that direct instantiation is not supported
     */
    private SpreadsheetFileDetector()
    {
        throw new UnsupportedOperationException("Instantiation not allowed");
    }

    /**
     * Detects XLS, XLSX, and ODS spreadsheet formats by examining file contents.
     * 
     * <p>
     * The detection process follows these steps:
     * <ol>
     * <li>Identifies legacy XLS files by checking for the OLE container signature and determining
     * the BIFF version of the contained workbook stream.</li>
     * <li>Identifies ZIP-based spreadsheet formats by inspecting their internal structures to
     * distinguish ODS from XLSX packages.</li>
     * <li>Returns an unknown format if no supported spreadsheet format is identified.</li>
     * </ol>
     *
     * @param fpath
     *        path to the spreadsheet file
     * @return the detection result, including the detected format, BIFF version, and workbook
     *         validation status
     * 
     * @throws IOException
     *         if an I/O error occurs while inspecting the file or validating the workbook
     */
    public static DetectionResult detect(Path fpath) throws IOException
    {
        if (fpath != null && Files.isRegularFile(fpath) && Files.isReadable(fpath))
        {
            if (containsMagicNumber(fpath, OLE_SIGNATURE))
            {
                BiffVersion version = detectBiffVersion(fpath);

                if (version != BiffVersion.UNKNOWN)
                {
                    return new DetectionResult(SpreadsheetFormat.XLS, version, isPoiValidWorkbook(fpath));
                }
            }

            else if (containsMagicNumber(fpath, ZIP_SIGNATURE))
            {
                if (isMagicNumberODS(fpath))
                {
                    /*
                     * The ODS structure is recognised, but its workbook
                     * contents are not validated by Apache POI.
                     */
                    return new DetectionResult(SpreadsheetFormat.ODS, BiffVersion.UNKNOWN, false);
                }

                if (isXlsxPackage(fpath))
                {
                    boolean valid = isPoiValidWorkbook(fpath);
                    return new DetectionResult(SpreadsheetFormat.XLSX, BiffVersion.UNKNOWN, valid);
                }
            }
        }

        return new DetectionResult(SpreadsheetFormat.UNKNOWN, BiffVersion.UNKNOWN, false);
    }

    /**
     * Checks whether a file begins with the specified byte signature.
     *
     * @param fpath
     *        path to the file
     * @param signature
     *        byte signature to compare with the file's initial bytes
     * @return {@code true} if the file begins with the specified signature, otherwise {@code false}
     * 
     * @throws IOException
     *         if an I/O error occurs while reading the file
     */
    private static boolean containsMagicNumber(Path fpath, byte[] signature) throws IOException
    {
        int totalRead = 0;
        byte[] buffer = new byte[signature.length];

        try (BufferedInputStream input = new BufferedInputStream(Files.newInputStream(fpath)))
        {
            // Read until the buffer is full or the end of the file is reached.
            while (totalRead < buffer.length)
            {
                int bytesRead = input.read(buffer, totalRead, buffer.length - totalRead);

                if (bytesRead == -1)
                {
                    return false;
                }

                totalRead += bytesRead;
            }
        }

        for (int i = 0; i < signature.length; i++)
        {
            if (buffer[i] != signature[i])
            {
                return false;
            }
        }

        return true;
    }

    /**
     * Identifies the BIFF version of an Excel workbook by examining the first record in its OLE
     * workbook stream.
     *
     * @param fpath
     *        path to the Excel file
     * @return the detected BIFF version, or {@link BiffVersion#UNKNOWN} if the workbook stream is
     *         missing or its BIFF version cannot be identified
     * 
     * @throws IOException
     *         if an I/O error occurs while reading the file
     */
    private static BiffVersion detectBiffVersion(Path fpath) throws IOException
    {
        String stream;

        try (POIFSFileSystem fs = new POIFSFileSystem(fpath.toFile(), true))
        {
            if (fs.getRoot().hasEntry("Workbook"))
            {
                stream = "Workbook";
            }

            else if (fs.getRoot().hasEntry("Book"))
            {
                stream = "Book";
            }

            else
            {
                return BiffVersion.UNKNOWN;
            }

            try (InputStream input = fs.getRoot().createDocumentInputStream(stream))
            {
                int idLow = input.read();
                int idHigh = input.read();
                int lenLow = input.read();
                int lenHigh = input.read();

                // Sanity check
                if (idLow < 0 || idHigh < 0 || lenLow < 0 || lenHigh < 0)
                {
                    return BiffVersion.UNKNOWN;
                }

                // Compute little-endian values
                int recordId = (idHigh << 8) | idLow;
                int payload = (lenHigh << 8) | lenLow;

                // Minimum payload length is 4 bytes
                if (payload < 4)
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
                        if (payload < 16)
                        {
                            return BiffVersion.UNKNOWN;
                        }

                        return detectBiff5Or8(input);

                    default:
                        return BiffVersion.UNKNOWN;
                }
            }
        }
    }

    /**
     * Reads the BOF version field following the four-byte record header.
     *
     * @param input
     *        starting stream positioned at the BOF record payload
     * @return the detected BIFF5 or BIFF8 version, or {@link BiffVersion#UNKNOWN} if the version
     *         field is truncated or unrecognised
     * 
     * @throws IOException
     *         if an I/O error occurs while reading the stream
     */
    private static BiffVersion detectBiff5Or8(InputStream input) throws IOException
    {
        int versionLow = input.read();
        int versionHigh = input.read();

        if (versionLow < 0 || versionHigh < 0)
        {
            return BiffVersion.UNKNOWN;
        }

        int version = (versionHigh << 8) | versionLow;

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
    private static boolean isMagicNumberODS(Path path)
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

        catch (IOException exc)
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

        catch (IOException exc)
        {
            return false;
        }
    }

    /**
     * Command-line example.
     */
    public static void main(String[] args)
    {
        Path path = Paths.get("src/io/test/Praveen_batch_list.xlsx");

        try
        {
            DetectionResult result = detect(path);

            System.out.println("File: " + path.getFileName());

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