package net.ssehub.kernel_haven.io.excel;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import net.ssehub.kernel_haven.util.FormatException;

/**
 * A simplistic reader for reading Excel documents in <tt>XLSX</tt> or <tt>XLS</tt> format.
 * @author El-Sharkawy
 *
 */
public class SimpleExcelReader implements Closeable {
    
    private Workbook wb;
    private boolean ignoreEmptyRows;
    
    /**
     * Default constructor for reading a XSLX document. Will also consider empty lines during reading.
     * @param inputFile An XLSX document, which shall be parsed.
     * @throws IOException if an error occurs while reading the data
     * @throws FormatException if the contents of the file cannot be parsed
     * @throws IllegalStateException If the workbook given is password protected
     */
    public SimpleExcelReader(File inputFile) throws IOException, IllegalStateException, FormatException {
        this(inputFile, false);
    }
    
    /**
     * Constructor for reading a XSLX document. The second parameter may be used to specify whether empty lines shall
     * be considered.
     * @param inputFile An XLSX document, which shall be parsed.
     * @param ignoreEmptyRows <tt>true</tt> empty rows will be skipped, <tt>false</tt> all lines will be read.
     * @throws IOException if an error occurs while reading the data
     * @throws FormatException if the contents of the file cannot be parsed
     * @throws IllegalStateException If the workbook given is password protected
     */
    public SimpleExcelReader(File inputFile, boolean ignoreEmptyRows) throws IOException, IllegalStateException,
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
     * Reads the complete content and calls the {@link #close()} method.
     * @return The sheets of the Excel document.
     * @throws IOException If an I/O error occurs during the {@link #close()} method.s
     */
    public List<ReadonlySheet> readAll() throws IOException {
        List<ReadonlySheet> result = new ArrayList<>();
        Iterator<org.apache.poi.ss.usermodel.Sheet> itr = wb.iterator();
        while (itr.hasNext()) {
            org.apache.poi.ss.usermodel.Sheet sheet = itr.next();
            result.add(readSheet(sheet));
        }
        
        close();
        
        return result;
    }
    
    /**
     * Reads the contents from the specified sheet.
     * 
     * @param index of the sheet number (0-based physical & logical)
     * @return Sheet at the provided index
     * @throws IllegalArgumentException if the index is out of range (index
     *            &lt; 0 || index &gt;= getNumberOfSheets()).
     */
    public ReadonlySheet readSheet(int index) {
        org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheetAt(index);
        
        return readSheet(sheet);
    }
    
    /**
     * Reads one sheet.
     * @param sheet The sheet of the Excel document to read.
     * @return The read sheet, should not be <tt>null</tt>.
     */
    private ReadonlySheet readSheet(org.apache.poi.ss.usermodel.Sheet sheet) {
        ReadonlySheet result = new ReadonlySheet(sheet.getSheetName());
        
        Iterator<Row> rowIterator = sheet.rowIterator();
        Deque<Integer> groupedRows = new ArrayDeque<>();
        int groupLevel = 0;
        int previousRow = -1;
        while (rowIterator.hasNext()) {
            List<Object> rowContents = new ArrayList<>();
            Row currentRow = rowIterator.next();
            int currentGroupLevel = currentRow.getOutlineLevel();
            
            if (currentGroupLevel > groupLevel) {
                // Current row is sub element of the row before
                groupedRows.addFirst(previousRow + 1);
            } else if (currentGroupLevel < groupLevel) {
                // Current row does not belong to the current row anymore, save last grouping
                Integer groupingStart = groupedRows.pollFirst();
                result.addRowGrouping(groupingStart, previousRow);
            }
            groupLevel = currentGroupLevel;
            
            Iterator<Cell> cellIterator = currentRow.iterator();
            boolean isEmpty = true;
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                Object value = null;
                switch (currentCell.getCellTypeEnum()) {
                case STRING:
                    value = currentCell.getStringCellValue();
                    break;
                case NUMERIC:
                    value = currentCell.getNumericCellValue();
                    break;
                case BOOLEAN:
                    value = currentCell.getBooleanCellValue();
                    break;
                case FORMULA:
                    value = currentCell.getStringCellValue();
                    break;
                default: 
                    value = currentCell.getStringCellValue();
                    break;
                }
                
                isEmpty &= value == null;
                rowContents.add(value);
            }
            
            if (!ignoreEmptyRows || !isEmpty) {
                result.addRow(rowContents.toArray());
            }
            previousRow++;
        }
        
        if (groupLevel > 0) {
            // Group ends at the last line
            Integer groupingStart = groupedRows.pollFirst();
            result.addRowGrouping(groupingStart, previousRow);
        }
        
        return result;
    }

    @Override
    public void close() throws IOException {
        wb.close();
    }

}
