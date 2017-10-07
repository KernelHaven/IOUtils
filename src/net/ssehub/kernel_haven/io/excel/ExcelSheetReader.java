package net.ssehub.kernel_haven.io.excel;

import java.io.IOException;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.ssehub.kernel_haven.util.io.ITableReader;

/**
 * A reader for a single sheet inside an excel file. Instances are created by {@link ExcelBook}s.
 * This class provides additional information about the grouping of rows, see {@link #getGroupedRows()}.
 *
 * @author El-Sharkawy
 * @author Adam
 */
public class ExcelSheetReader implements ITableReader {

    private Sheet sheet;
    
    private String sheetName;
    
    private List<String[]> contents;

    private List<Group> groupedRows;
    
    private boolean ignoreEmptyRows;
    
    /**
     * The current position for {@link #readNextRow()}. Reset when {@link #close()} is called.
     */
    private Iterator<String[]> iterator;
    
    ExcelSheetReader(Sheet sheet, boolean ignoreEmptyRows) {
        this.sheet = sheet;
        this.sheetName = sheet.getSheetName();
        this.ignoreEmptyRows = ignoreEmptyRows;
        this.contents = new ArrayList<>();
        this.groupedRows = new ArrayList<>();
        
        read();
        
        this.iterator = contents.iterator(); 
    }
    
    private void read() {
        Iterator<Row> rowIterator = sheet.rowIterator();
        Deque<Integer> groupedRows = new ArrayDeque<>();
        int groupLevel = 0;
        int previousRow = -1;
        while (rowIterator.hasNext()) {
            List<String> rowContents = new ArrayList<>();
            Row currentRow = rowIterator.next();
            int currentGroupLevel = currentRow.getOutlineLevel();
            
            if (currentGroupLevel != groupLevel) {
                while (currentGroupLevel > groupLevel) {
                    // Current row is sub element of the row before
                    groupedRows.addFirst(previousRow + 1);
                    groupLevel++;
                }
                while (currentGroupLevel < groupLevel) {
                    // Current row does not belong to the current row anymore, save last grouping
                    Integer groupingStart = groupedRows.pollFirst();
                    this.groupedRows.add(new Group(groupingStart, previousRow));
                    groupLevel--;
                }
            }
            
//            if (currentGroupLevel > groupLevel) {
//            } else if (currentGroupLevel < groupLevel) {
//                // Current row does not belong to the current row anymore, save last grouping
//                Integer groupingStart = groupedRows.pollFirst();
//                result.addRowGrouping(groupingStart, previousRow);
//            }
//            groupLevel = currentGroupLevel;
            
            Iterator<Cell> cellIterator = currentRow.iterator();
            boolean isEmpty = true;
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                String value = null;
                switch (currentCell.getCellTypeEnum()) {
                case STRING:
                    value = currentCell.getStringCellValue();
                    break;
                case NUMERIC:
                    value = currentCell.getNumericCellValue() + "";
                    break;
                case BOOLEAN:
                    value = currentCell.getBooleanCellValue() + "";
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
                this.contents.add(rowContents.toArray(new String[0]));
            }
            previousRow++;
        }
        
        while (groupLevel > 0) {
            // Group ends at the last line
            Integer groupingStart = groupedRows.pollFirst();
            int lastRow = Math.min(previousRow, this.contents.size() - 1);
            this.groupedRows.add(new Group(groupingStart, lastRow));
            groupLevel--;
        }
    }
    
    public String getSheetName() {
        return sheetName;
    }
    
    public List<Group> getGroupedRows() {
        return groupedRows;
    }
    
    @Override
    public void close() {
        // no need to close anything, just reset the iterator
        iterator = contents.iterator();
    }

    @Override
    public String[] readNextRow() throws IOException {
        String[] result;
        if (iterator.hasNext()) {
            result = iterator.next();
        } else {
            result = null;
        }
        return result;
    }

}
