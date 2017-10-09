package net.ssehub.kernel_haven.io.excel;

import java.io.BufferedOutputStream;
import java.io.Closeable;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import net.ssehub.kernel_haven.util.FormatException;
import net.ssehub.kernel_haven.util.io.ITableCollection;

/**
 * A wrapper around an excel book. A book contains several sheets. The individual sheets can be accessed through
 * {@link ExcelSheetReader}s, see {@link #getAllSheetReaders()} and {@link #getSheetReader(int)}.
 *
 * @author Adam
 * @author El-Sharkawy
 */
public class ExcelBook implements ITableCollection, Closeable {
    
    private static enum Mode {
        // Add further state if needed
        
        // Existing workbook, which shall not be changed
        READ_ONLY,
        
        // New workbook, read (temporary) data and add new data
        WRITE_NEW_WB;
    }
    
    private Workbook wb;
    
    private boolean ignoreEmptyRows;
    private Mode mode;
    private File destinationFile;
    
    /**
     * Default constructor for reading and writing a Excel documents (XLSX, XLS).
     * Will also consider empty lines during reading.
     * 
     * @param destinationFile An Excel document, which shall be parsed (if existing) or be written (if not existing).
     * 
     * @throws IOException if an error occurs while reading the data
     * @throws FormatException if the contents of the file cannot be parsed
     * @throws IllegalStateException If the workbook given is password protected
     */
    public ExcelBook(File destinationFile) throws IOException, IllegalStateException, FormatException {
        this(destinationFile, false);
    }
    
    /**
     * Constructor for reading and writing a Excel documents (XLSX, XLS).
     * The second parameter may be used to specify whether empty lines shall be considered.
     * 
     * @param destinationFile An Excel document, which shall be parsed (if existing) or be written (if not existing).
     * @param ignoreEmptyRows <tt>true</tt> empty rows will be skipped, <tt>false</tt> all lines will be read.
     * 
     * @throws IOException if an error occurs while reading the data
     * @throws FormatException if the contents of the file cannot be parsed
     * @throws IllegalStateException If the workbook given is password protected
     */
    public ExcelBook(File destinationFile, boolean ignoreEmptyRows) throws IOException, IllegalStateException,
        FormatException {
        
        this.ignoreEmptyRows = ignoreEmptyRows;
        this.destinationFile = destinationFile;
        if (!destinationFile.exists()) {
            if (destinationFile.createNewFile()) {
                mode = Mode.WRITE_NEW_WB;
                wb = new XSSFWorkbook();
            } else {
                throw new IOException("Specified file does not exist and could not be created: "
                    + destinationFile.getAbsolutePath());
            }
        } else {
            try {
                mode = Mode.READ_ONLY;
                /* Using a File object allows for lower memory consumption, while an InputStream requires more memory
                 * as it has to buffer the whole file.
                 */
                wb = WorkbookFactory.create(destinationFile);
            } catch (InvalidFormatException e) {
                throw new FormatException(e);
            }
        }
    }
    
    /**
     * Returns {@link ExcelSheetReader}s for all sheets in this book.
     * 
     * @return Readers for all sheets of the Excel document.
     */
    public List<ExcelSheetReader> getAllSheetReaders() {
        List<ExcelSheetReader> result = new ArrayList<>();
        
        for (Sheet sheet : wb) {
            result.add(new ExcelSheetReader(sheet, ignoreEmptyRows));
        }
        
        return result;
    }
    
    /**
     * Returns a reader for the specified sheet.
     * 
     * @param index Index of the sheet number (0-based physical & logical)
     * @return A reader for the sheet at the provided index.
     * 
     * @throws IllegalArgumentException if the index is out of range (index
     *            &lt; 0 || index &gt;= getNumberOfSheets()).
     */
    public ExcelSheetReader getReader(int index) {
        Sheet sheet = wb.getSheetAt(index);
        
        return new ExcelSheetReader(sheet, ignoreEmptyRows);
    }
    
    @Override
    public Set<String> getTableNames() throws IOException {
        Set<String> result = new HashSet<>();
        
        for (Sheet sheet : wb) {
            result.add(sheet.getSheetName());
        }
        
        return result;
    }
    
    @Override
    public ExcelSheetReader getReader(String name) {
        ExcelSheetReader result = null;
        for (Sheet sheet : wb) {
            if (sheet.getSheetName().equals(name)) {
                result = new ExcelSheetReader(sheet, ignoreEmptyRows);
                break;
            }
        }
        return result;
    }

    @Override
    public ExcelSheetWriter getWriter(String name) throws IOException {
        switch (mode) {
        case READ_ONLY:
            throw new UnsupportedOperationException("Sheet was oppened in read only mode: "
                + destinationFile.getAbsolutePath());
        case WRITE_NEW_WB:
            // falls through
        default:
            String safeName = WorkbookUtil.createSafeSheetName(name);
            IllegalArgumentException exception = null;
            Sheet sheet = null;
            try {
                sheet = wb.createSheet(safeName);
            } catch (IllegalArgumentException exc) {
                exception = exc;
                byte id = 0;
                while (null == sheet && id < Byte.MAX_VALUE) {
                    String tmpName = WorkbookUtil.createSafeSheetName(safeName + id);
                    try {
                        sheet = wb.createSheet(tmpName);
                        /* Add sheets at the front by default
                         * This is done to show final results at the beginning of document and intermediate results
                         * at the end of document.
                         */
                        wb.setSheetOrder(sheet.getSheetName(), 0);
                    } catch (IllegalArgumentException exc2) {
                        // No action needed
                    }
                    id++;
                }
            }
            if (null == sheet) {
                throw new IOException("Could not create sheet \"" + safeName + "\", cause: " + exception.getMessage());
            }
            return new ExcelSheetWriter(sheet);
        }
    }
    
    @Override
    public void close() throws IOException {
        switch (mode) {
        case WRITE_NEW_WB:
            BufferedOutputStream fileOut = new BufferedOutputStream(new FileOutputStream(destinationFile));
            wb.write(fileOut);
        case READ_ONLY:
            break;
        default:
            throw new IllegalStateException("Unexpected close opperation for state: " + mode.name());
        }
        wb.close();
    }

}
