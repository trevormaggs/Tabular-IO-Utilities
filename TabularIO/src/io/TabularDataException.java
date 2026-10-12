package io;

import java.util.Objects;

/**
 * Represents a failure during tabular data or workbook read/write operations. Provides an
 * {@link ErrorCode} to categorise the failure and supports retaining an underlying cause for
 * diagnostic purposes.
 *
 * @author Trevor Maggs
 * @version 0.1
 * @since 7 October 2026
 */
public class TabularDataException extends Exception
{
    private static final long serialVersionUID = 1L;

    /**
     * Categorises the type of failure encountered during tabular data processing.
     */
    public enum ErrorCode
    {
        /** An unspecified or uncategorised failure. */
        GENERAL_ERROR,

        /** Access to the file was denied. */
        FILE_ACCESS_DENIED,

        /** The requested file could not be found. */
        FILE_NOT_FOUND,

        /** The file format is not supported. */
        UNSUPPORTED_FORMAT,

        /** Reading from an input stream failed. */
        STREAM_READ_FAILED,

        /** Writing to an output stream failed. */
        STREAM_WRITE_FAILED,

        /** Parsing the tabular data failed. */
        PARSE_FAILED,

        /** The requested sheet could not be found or identified. */
        SHEET_UNKNOWN
    }

    private final ErrorCode errorCode;

    /**
     * Constructs a {@code TabularDataException} with a general error code and no message.
     */
    public TabularDataException()
    {
        super();
        this.errorCode = ErrorCode.GENERAL_ERROR;
    }

    /**
     * Constructs a {@code TabularDataException} with the specified detail message.
     *
     * @param message
     *        the detail message
     */
    public TabularDataException(String message)
    {
        super(message);
        this.errorCode = ErrorCode.GENERAL_ERROR;
    }

    /**
     * Constructs a {@code TabularDataException} with the specified cause.
     *
     * @param cause
     *        the underlying cause
     */
    public TabularDataException(Throwable cause)
    {
        super(cause);
        this.errorCode = ErrorCode.GENERAL_ERROR;
    }

    /**
     * Constructs a {@code TabularDataException} with a detail message and underlying cause.
     *
     * @param message
     *        the detail message
     * @param cause
     *        the underlying cause
     */
    public TabularDataException(String message, Throwable cause)
    {
        super(message, cause);
        this.errorCode = ErrorCode.GENERAL_ERROR;
    }

    /**
     * Constructs a {@code TabularDataException} with a custom error code and detail message.
     *
     * @param errorCode
     *        the error code categorisation
     * @param message
     *        the detail message
     */
    public TabularDataException(ErrorCode errorCode, String message)
    {
        super(message);
        this.errorCode = errorCode;
    }

    /**
     * Constructs a {@code TabularDataException} with a custom error code, detail message, and
     * cause.
     *
     * @param errorCode
     *        the error code categorisation
     * @param message
     *        the detail message
     * @param cause
     *        the underlying cause
     */
    public TabularDataException(ErrorCode errorCode, String message, Throwable cause)
    {
        super(message, cause);
        this.errorCode = Objects.requireNonNull(errorCode, "errorCode");
    }

    /**
     * Returns the specific error code associated with this exception.
     *
     * @return the {@link ErrorCode}
     */
    public ErrorCode getErrorCode()
    {
        return errorCode;
    }
}