# Tabular I/O Utilities

## Overview

**Tabular I/O Utilities** is a reusable Java library for reading, writing, and managing tabular data across different projects.

The library supports three file formats:

- **CSV (`.csv`)** - comma-separated values.
- **Excel Workbook (`.xlsx`)** - Office Open XML spreadsheet format.
- **Excel 97–2003 Workbook (`.xls`)** - legacy Microsoft Excel binary workbook format.

Designed for integration into Java applications, this library provides common interfaces for reading and writing structured data, alongside spreadsheet-specific features such as multiple worksheets, automatic filters, and alternating row styles.

## Documentation

Complete API documentation is available in the published Javadoc:

[**View Javadoc Documentation**](https://trevormaggs.github.io/Tabular_IO_Utilities/)

## Features

### CSV and Tabular Data Processing

The `Readable` and `Writable` interfaces define common contracts for reading and writing tabular data.

Features include:

- Opening and closing data resources.
- Retrieving file paths and data dimensions.
- Iterating over rows represented as arrays of strings.
- Inserting text and numeric values.
- Writing column headers and row content.
- Locating columns by cell content or regular expressions.

### Spreadsheet Reading

The `ExcelReadable` interface extends `Readable` with workbook-specific operations:

- Retrieving the active worksheet name.
- Obtaining the total number of worksheets.
- Reading worksheets by name or zero-based index.

### Spreadsheet Writing and Formatting

The `ExcelWritable` interface extends `Writable` with additional functionality:

- Creating worksheets.
- Retrieving the current worksheet name.
- Configuring automatic filters.
- Applying alternating row styles.
- Accessing the alternating row style factory.

### Exception Handling

The `TabularDataException` class provides a dedicated exception type for tabular data operations.

Its `ErrorCode` enumeration categorises errors such as file access failures, unsupported formats, stream read/write failures, parsing errors, and worksheet-related problems.

## Design

The library uses interfaces to establish consistent contracts for reading and writing structured data.

| Interface or class | Purpose |
|---|---|
| `Readable` | General tabular data reading operations. |
| `Writable` | General tabular data writing operations. |
| `ExcelReadable` | Spreadsheet reading, including worksheet selection. |
| `ExcelWritable` | Spreadsheet writing and formatting operations. |
| `TabularDataException` | Exception handling and error categorisation. |

Spreadsheet-specific styling support is provided by the `io.style` package, which includes `AlternatingRowFactory` and `ColourStyle`.

This separation allows applications to work with tabular data through common interfaces while retaining access to spreadsheet-specific features.

## Technologies

- **Java** - core programming language.
- **Eclipse IDE** - development environment.
- **Apache POI 5.5.1** - reading and writing Microsoft Excel workbooks in `.xls` and `.xlsx` formats.

## Java Compatibility

- Minimum Java version: Core Java 8 (JDK 8).
- The codebase favours traditional Java syntax, including anonymous inner classes instead of lambda expressions, for readability and consistency.

## Dependencies

The library uses [Apache POI](https://poi.apache.org/) version **5.5.1** to read and write Microsoft Excel workbooks in `.xls` and `.xlsx` formats.

Apache POI requires additional libraries at runtime. When importing the project into Eclipse, ensure that Apache POI and its required dependencies are available on the Java Build Path.

The libraries used for development are maintained locally in the Eclipse project and are not included in the GitHub repository. Users setting up the project must obtain the appropriate libraries separately.

CSV processing uses a text-based approach and does not require Apache POI for basic reading and writing operations.


## Usage

The library can be integrated into Java applications through its reading and writing interfaces and their concrete implementations. The examples below demonstrate common ways to read CSV files, read Excel workbooks, stream large `.xlsx` files, and write formatted workbooks.

Each example is collapsible. Expand a section to view its code.

### Usage Examples

<details>
<summary><strong>1. Reading CSV Files</strong></summary>

For lightweight, text-based tabular processing without external dependencies, use the CSV reader implementation.

```java
Path csvFile = Paths.get("employees.csv");

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
```
</details>

<details>
<summary><strong>2. Reading CSV Files via an Iterator</strong></summary>

Alternatively, you can implement an Iterator to access to the data. It is also lightweight and fast.

```java
Path csvFile = Paths.get("metadata.csv");

try (Readable reader = new ReadCSV(csvFile))
{
	System.out.printf("Dimension: Rows[%d], Columns [%d]%n%n", reader.getRowCount(), reader.getColumnCount());
	{	
		Iterator<String[]> iterator = reader.iterator();
		
		// Uses an Iterator provided to loop over each row as a String array
		while (iterator.hasNext())
		{
			String[] row = iterator.next();
			System.out.println(Arrays.toString(row));
		}
	}
}

catch (TabularDataException exc)
{
  System.err.printf("[%s] %s%n", exc.getErrorCode(), exc.getMessage());
}
```
</details>

<details>
<summary><strong>3. Reading CSV Files using a callback Consumer pattern</strong></summary>

You can also opt to apply a callback mechanism, where you can define your desired logic and supply to the constructor.

```java
Path csvFile = Paths.get("customer.csv");

BiConsumer<String[], Integer> consumer = new BiConsumer<String[], Integer>()
{
	@Override
	public void accept(String[] row, Integer lineNumber)
	{
		if (lineNumber == 0)
		{		
			// Process headers if required
			System.out.println(Arrays.toString(row));
		}
		
		else
		{
			// Process regular contents
			System.out.println(Arrays.toString(row));
		}
	}
};

try (ReadCSV reader = new ReadCSV(csvFile, consumer))
{
  // Parsing and processing are fully completed inside the try-with-resources block.
  System.out.printf("\nDimension: Rows[%d], Columns [%d]%n%n", reader.getRowCount(), reader.getColumnCount());
}

catch (TabularDataException exc)
{
  System.err.printf("[%s] %s%n", exc.getErrorCode(), exc.getMessage());
}
```
</details>

<details>
<summary><strong>4. Reading Standard Excel Workbooks via an Iterator (<code>ReadWorkbook</code>)</strong></summary>

Use `ReadWorkbook` when you want to iterate every row within the chosen workbook, including formula evaluation, and formatted strings.

```java
ReadWorkbook book = new ReadWorkbook(Paths.get("data", "financial_report.xlsx"), "Summary");

try
{
    book.open();

    Iterator<String[]> it = book.iterator();
    
    System.out.printf("%s%n", book.getFilePath());
    System.out.printf("%s%n", book.getWorksheetName());  

    while (it.hasNext())
    {
        System.out.println(Arrays.toString(it.next()));
    }
}

catch (Exception exc)
{
  System.err.println(exc.getMessage());
}

finally
{
    book.close();
}
```

In this example:

- **Resource management:** The workbook is opened using `open()` and closed explicitly using `close()`. When not using a try-with-resources statement, ensure that `close()` is called even if an exception occurs during processing.
- **Worksheet context:** `book.getWorksheetName()` returns the currently selected worksheet name.

</details>        
<details>
<summary><strong>5. Reading Standard Excel Workbooks (<code>ReadWorkbook</code>)</strong></summary>

Use `ReadWorkbook` when you want workbook access, formula evaluation, and formatted strings.

```java
Path excelFile = Paths.get("data", "financial_report.xlsx");

try (ExcelReadable reader = new ReadWorkbook(excelFile, true, "Summary"))
{
    reader.open();
    System.out.printf("Active Worksheet: %s%n", reader.getWorksheetName());
    System.out.printf("Dimensions - Rows: %d, Columns: %d%n", reader.getRowCount(), reader.getColumnCount());

    // Locate a specific column index by header title
    int headerRow = 0;
    int colIndex = reader.findColumnIndexByContent("Department", headerRow);
    System.out.printf("Found 'Department' at column index: %d%n", colIndex);

    // Loop over rows using standard iteration
    for (String[] row : reader)
    {
        System.out.println(Arrays.toString(row));
    }
}
catch (TabularDataException exc)
{
    System.err.printf("Workbook read error [%s]: %s%n", exc.getErrorCode(), exc.getMessage());
}
```
</details>
<details>
<summary><strong>6. Streaming Large Excel Files (<code>EventExcelReader</code>)</strong></summary>

For large `.xlsx` files, `EventExcelReader` supports low-memory, event-driven SAX parsing using a consumer callback pattern.

```java
Path largeFile = Paths.get("data", "massive_transaction_log.xlsx");

// Create reader with style support and deferred worksheet selection
try (EventExcelReader reader = new EventExcelReader(largeFile, true))
{
    reader.open();

    // Stream rows incrementally from the "Transactions" worksheet using an anonymous inner-class consumer
    reader.readWorksheet("Transactions", new Consumer<String[]>()
    {
        @Override
        public void accept(String[] row)
        {
            // Do whatever you want to do with each row as it is delivered here
            if (row.length > 2 && "FAILED".equalsIgnoreCase(row[2]))
            {
                System.out.println("Alert - Failed Transaction: " + Arrays.toString(row));
            }
        }
    });
}
catch (TabularDataException exc)
{
    System.err.printf("Streaming error [%s]: %s%n", exc.getErrorCode(), exc.getMessage());
}
```
</details>
<details>
<summary><strong>7. Reading Excel Rows as Maps (<code>NamedRowConsumer</code>)</strong></summary>

When processing worksheets with known column headers, `NamedRowConsumer` provides each row as a `Map<String, String>`. This allows values to be accessed by column name instead of by numeric index. You can specify which columns to include when constructing the consumer.

The following example reads the `Owner name` and `group` columns from the `Sheet3` worksheet and processes each row using its header names.

```java
Path excelFile = Paths.get("data", "workbook.xlsx");

try (EventExcelReader reader = new EventExcelReader(excelFile, true))
{
    reader.open();

    System.out.printf("Worksheet Name: %s%n", reader.getWorksheetName());

    reader.readWorksheet("Sheet3", new NamedRowConsumer("Owner name", "group")
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
```

In this example:

- **Column selection:** Only `Owner name` and `group` are included in each row map. If no column names are specified, all columns will be read.
- **Header-based access:** Values are retrieved using `rowMap.get("Owner name")` and `rowMap.get("group")`.
- **Row indexes:** `rowIndex` is a zero-based index of data rows. The first data row has index `0`, regardless of whether a header row is present.
- **Worksheet context:** `reader.getWorksheetName()` returns the currently selected worksheet name.

This approach is useful when column positions may change but the header names remain consistent.
</details>
<details>
<summary><strong>8. Writing Excel Workbooks (<code>ExcelWritable</code>)</strong></summary>

Create a workbook, write headers and data rows, and configure an auto-filter.

```java
Path outputFile = Paths.get("output", "generated_report.xlsx");

try (ExcelWritable writer = new ExcelWorkbookWriter(outputFile))
{
    writer.open();

    // Create a worksheet and write headers
    writer.createWorksheet("Report Data");
    writer.writeRow(new String[] { "ID", "Name", "Status", "Amount" });

    // Write content rows
    writer.writeRow(new String[] { "101", "Acme Corp", "Active", "1500.00" });
    writer.writeRow(new String[] { "102", "Beta LLC", "Pending", "450.50" });

    // Configure auto-filters across headers
    writer.setAutoFilter(0, 0, 1, 3);
}
catch (TabularDataException exc)
{
    System.err.printf("Workbook write error [%s]: %s%n", exc.getErrorCode(), exc.getMessage());
}
```
</details>

For API details and additional operations, refer to the [published Javadoc](https://trevormaggs.github.io/Tabular_IO_Utilities/).

## Project Status

The library provides a foundation for reusable CSV and spreadsheet data processing in Java applications.

Future development may include support for OpenDocument Spreadsheet (`.ods`) format used by LibreOffice Calc, extending compatibility beyond CSV and Microsoft Excel workbooks and further documentation of implementation details and usage examples.

## Credits

Developed and maintained by **Trevor Maggs**.

Bug reports, suggestions, and feedback are welcome.

**Email:** [trevmaggs@tpg.com.au](mailto:trevmaggs@tpg.com.au)

## License

This project is nearing completion, with further testing currently underway.

An open-source licence will be added shortly. Until a licence is published, the terms under which this library may be used, modified, or redistributed have not yet been formally established.
