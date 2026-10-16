package io.test;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Iterator;
import java.util.Map;
import io.reader.EventExcelReader;
import io.reader.ReadCSV;
import io.reader.ReadWorkbook;
import io.style.AlternatingRowFactory;
import io.writer.WriteWorkbook;
import io.ExcelWritable;
import io.Readable;
import io.TabularDataException;

public class TestWorkbook
{
    public static void testReadWorkbook()
    {
        try
        {
            ReadWorkbook book = new ReadWorkbook(Paths.get("src\\io\\test\\TurramurraHomeConstruction.xlsx"), "Projected costs");
            book.open();
            // book.setWorksheet("tab2");

            Iterator<String[]> it = book.iterator();

            System.out.printf("%s%n", book.getFilePath());
            System.out.printf("%s%n", book.getWorksheetName());

            System.out.printf("%s%n", book.findColumnIndexByContent("Amount"));
            System.out.printf("%s%n", book.findColumnIndexByContent("Category"));
            System.out.printf("%s%n", book.findColumnIndexByContent("Owner name"));

            while (it.hasNext())
            {
                System.out.printf("%s%n", Arrays.toString(it.next()));
            }

            book.close();
        }

        catch (Exception exc)
        {
            System.err.println(exc.getMessage());
        }
    }

    public static void testAlternatingRowFactory()
    {
        try
        {
            ExcelWritable book = new WriteWorkbook("gem.XLSx", "Testing");

            String[] header = {
                    "LittleFix Version", "Policy File", "Run Date",
                    "Hostname", "Check Name", "Check Description",
                    "Setting Type", "Expected Setting", "Current Setting",
                    "Result", "Severity"};

            String[] content = {
                    "Trevor", "Nina", "Gemma",
                    "Misty", "Wesley", "Adrian",
                    "Savannah", "Lincoln", "Elizabeth",
                    "Abraham"};

            int min = 50;
            int max = 100;

            book.open();

            System.out.printf("%s%n", book.getFilePath());
            System.out.printf("%s%n", book.getWorksheetName());

            book.writeHeader(header);

            for (String str : content)
            {
                book.addRow();
                book.insertCell(Math.floor(Math.random() * (max - min + 1) + min));
                book.insertCell(str);
            }

            book.setAutoFilter(0, 0, content.length - 1, header.length - 1);
            // book.alternateRowStyles(0, 0, content.length - 1, header.length - 1,
            // ColourStyle.ORANGE);

            AlternatingRowFactory factory = book.getInstanceAlternatingRow();
            factory.setHeaderRowRGB(192, 80, 77).setEvenRowRGB(204, 255, 204).setOddRowRGB(128, 100, 162).setBorderRGB(128, 128, 128).apply(0, 0, content.length - 1, header.length - 1);

            book.close();
        }

        catch (Exception exc)
        {
            exc.printStackTrace();
        }
    }

    public static void testReadCSV()
    {
        Path csvFile = Paths.get("src/io/test/DESKTOP-E0REK70_metadata.csv");

        try (Readable reader = new ReadCSV(csvFile))
        {
            reader.open();

            System.out.printf("Dimension: Rows[%d], Columns [%d]%n%n", reader.getRowCount(), reader.getColumnCount());

            // Iterate over rows represented as String arrays
            for (String[] row : reader)
            {
                System.out.println(Arrays.toString(row));
            }
        }

        catch (TabularDataException exc)
        {
            System.err.printf("[%s] %s%n", exc.getErrorCode(), exc.getMessage());
        }
    }

    public static void testEventExcelReader()
    {
        Path excelFile = Paths.get("src/io/test/Praveen_batch_list.xlsx");

        try (EventExcelReader reader = new EventExcelReader(excelFile, true))
        {
            reader.open();

            System.out.printf("Worksheet Name: %s%n", reader.getWorksheetName());   

            reader.readWorksheet("CDW", new EventExcelReader.NamedRowConsumer("Owner name", "group")
            {
                @Override
                public void acceptRow(Map<String, String> rowMap, int rowIndex)
                {
                    String owner = rowMap.get("Owner name");
                    String group = rowMap.get("group");

                    System.out.printf("Row %d -> Name: %s\tGroup: %s%n", rowIndex, owner, group);
                }
            });
        }

        catch (TabularDataException exc)
        {
            System.err.printf("Streaming error [%s]: %s%n", exc.getErrorCode(), exc.getMessage());
        }
    }

    public static void main(String[] args)
    {
        // testReadWorkbook();
        // testAlternatingRowFactory();
        // testReadCSV();
        testEventExcelReader();
    }
}