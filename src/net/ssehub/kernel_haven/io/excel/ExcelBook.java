package net.ssehub.kernel_haven.io.excel;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import net.ssehub.kernel_haven.util.FormatException;
import net.ssehub.kernel_haven.util.io.ITableCollection;
import net.ssehub.kernel_haven.util.io.ITableWriter;

/**
 * A wrapper around an excel book. A book contains several sheets. The individual sheets can be accessed through
 * {@link ExcelSheetReader}s, see {@link #getAllSheetReaders()} and {@link #getSheetReader(int)}.
 *
 * @author Adam
 * @author El-Sharkawy
 */
public class ExcelBook implements ITableCollection, Closeable {
    
    private Workbook wb;
    
    private boolean ignoreEmptyRows;
    
    /**
     * Default constructor for reading a XSLX document. Will also consider empty lines during reading.
     * 
     * @param inputFile An XLSX document, which shall be parsed.
     * 
     * @throws IOException if an error occurs while reading the data
     * @throws FormatException if the contents of the file cannot be parsed
     * @throws IllegalStateException If the workbook given is password protected
     */
    public ExcelBook(File inputFile) throws IOException, IllegalStateException, FormatException {
        this(inputFile, false);
    }
    
    /**
     * Constructor for reading a XSLX document. The second parameter may be used to specify whether empty lines shall
     * be considered.
     * 
     * @param inputFile An XLSX document, which shall be parsed.
     * @param ignoreEmptyRows <tt>true</tt> empty rows will be skipped, <tt>false</tt> all lines will be read.
     * 
     * @throws IOException if an error occurs while reading the data
     * @throws FormatException if the contents of the file cannot be parsed
     * @throws IllegalStateException If the workbook given is password protected
     */
    public ExcelBook(File inputFile, boolean ignoreEmptyRows) throws IOException, IllegalStateException,
        FormatException {
        
        this.ignoreEmptyRows = ignoreEmptyRows;
        if (!inputFile.exists()) {
            throw new IOException(inputFile.getAbsolutePath() + " does not exist.");
        }
        try {
            wb = WorkbookFactory.create(inputFile);
        } catch (InvalidFormatException e) {
            throw new FormatException(e);
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
    public ITableWriter getWriter(String name) throws IOException {
        throw new UnsupportedOperationException("Writing excel sheets is not yet supported");
    }
    
    @Override
    public void close() throws IOException {
        wb.close();
    }

}
