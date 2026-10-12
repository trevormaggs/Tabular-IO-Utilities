# Tabular I/O Utilities

## Overview

**Tabular I/O Utilities** is a reusable Java library for reading, writing, and managing tabular data across different projects.

The library supports three file formats:

- **CSV (`.csv`)** — comma-separated values.
- **Excel Workbook (`.xlsx`)** — Office Open XML spreadsheet format.
- **Excel 97–2003 Workbook (`.xls`)** — legacy Microsoft Excel binary workbook format.

Designed for integration into different Java applications, the library provides a common interface for reading and writing structured data, alongside spreadsheet-specific features such as multiple worksheets, automatic filters, and alternating row styles.

Future development is intended to include support for the OpenDocument Spreadsheet (`.ods`) format used by LibreOffice Calc, extending the library's compatibility beyond CSV and Microsoft Excel workbooks.

## Documentation

The complete API documentation is available in the published Javadoc:

**[View Javadoc Documentation](https://trevormaggs.github.io/Tabular_IO_Utilities/)**

## Features

### CSV and tabular data processing

The `Readable` and `Writable` interfaces define common contracts for reading and writing tabular data.

Features include:

- Opening and closing data resources.
- Retrieving file paths and data dimensions.
- Iterating over rows represented as arrays of strings.
- Inserting text and numeric values.
- Writing column headers and row content.
- Locating columns by cell content or regular expressions.

### Spreadsheet reading

The `ExcelReadable` interface extends `Readable` with workbook-specific operations:

- Retrieving the active worksheet name.
- Obtaining the total number of worksheets.
- Reading worksheets by name or zero-based index.

### Spreadsheet writing and formatting

The `ExcelWritable` interface extends `Writable` with additional functionality:

- Creating worksheets.
- Retrieving the current worksheet name.
- Configuring automatic filters.
- Applying alternating row styles.
- Accessing the alternating row style factory.

### Exception handling

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

This separation allows applications to work with tabular data through common interfaces while retaining access to features specific to spreadsheet formats.

## Technologies

- **Java** — core programming language.
- **Apache POI** — library used to read and write Microsoft Excel workbooks in `.xls` and `.xlsx` formats.
- **Eclipse IDE** — development environment.
- **Java `AutoCloseable`** — resource lifecycle management.
- **Java regular expressions** — column-content searching.

## Java Compatibility

- **Minimum Java version:** Java 8 (JDK 8).
- The codebase is written without lambda expressions, favouring traditional Java syntax for compatibility and readability.

## Dependencies

This project uses the [Apache POI](https://poi.apache.org/) library to provide support for Microsoft Excel workbook formats.

Apache POI supports both the legacy `.xls` format and the modern `.xlsx` format. CSV processing uses a text-based tabular data approach.

Ensure that the required Apache POI libraries and their dependencies are available on the Java Build Path when importing the project into Eclipse.

### Usage

The library is intended to be integrated into Java applications through its reading and writing interfaces and their concrete implementations.

Usage examples, dependency requirements, and supported format details can be added alongside the corresponding implementations.

## Project status

The library provides a foundation for reusable CSV and spreadsheet data processing in Java applications.

Further documentation can describe the supported file formats, implementation details, dependencies, and usage examples as the project develops.

## Credits

Developed and maintained by **Trevor Maggs**.

Bug reports, suggestions and feedback are welcome.

**Email:** [trevmaggs@tpg.com.au](mailto:trevmaggs@tpg.com.au)

## License

This project is nearing completion, with further testing currently underway.

An open-source `licence` will be added shortly. Until a licence is published, the terms under which this library may be used, modified, or redistributed have not yet been formally established.
