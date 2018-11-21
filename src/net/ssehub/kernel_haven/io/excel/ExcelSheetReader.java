package net.ssehub.kernel_haven.io.excel;

import static net.ssehub.kernel_haven.util.null_checks.NullHelpers.notNull;

import java.io.IOException;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Deque;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.ssehub.kernel_haven.util.io.ITableReader;
import net.ssehub.kernel_haven.util.null_checks.NonNull;
import net.ssehub.kernel_haven.util.null_checks.Nullable;

/**
 * A reader for a single sheet inside an excel file. Instances are created by {@link ExcelBook}s.
 * This class provides additional information about the grouping of rows, see {@link #getGroupedRows()}.
 *
 * @author El-Sharkawy
 * @author Adam
 */
public class ExcelSheetReader implements ITableReader {

    private @NonNull Sheet sheet;
    
    private @NonNull String sheetName;
    
    /**
     * Caches the result of {@link #getGroupedRows()}. <code>null</code> if not yet called.
     */
    private @Nullable List<@NonNull Group> groupedRows;
    
    private boolean ignoreEmptyRows;
    
    /**
     * Iterator for the rows in the sheet. Reset when {@link #close()} is called.
     */
    private @NonNull Iterator<Row> rowIterator;
    
    /**
     * The number of columns we expect. (read from the first row)
     */
    private int nColumns;
    
    /**
     * The current row number.
     */
    private int currentRowNumber;
    
    /**
     * Creates an reader for the given sheet.
     * 
     * @param sheet The sheet to create this reader for.
     * @param ignoreEmptyRows Whether empty rows should be ignored or not.
     */
    ExcelSheetReader(@NonNull Sheet sheet, boolean ignoreEmptyRows) {
        this.sheet = sheet;
        this.sheetName = notNull(sheet.getSheetName());
        this.ignoreEmptyRows = ignoreEmptyRows;
        
        this.nColumns = 0;
        if (sheet.getRow(0) != null) {
            this.nColumns = sheet.getRow(0).getLastCellNum();
        }
        this.rowIterator = notNull(sheet.rowIterator());
    }
    
    /**
     * Returns the name of this sheet.
     * 
     * @return The name of this sheet.
     */
    public @NonNull String getSheetName() {
        return sheetName;
    }
    
    /**
     * Returns a list of row groupings of this sheet.
     * 
     * @return A list containing all {@link Group}s of rows in this sheet.
     */
    public @NonNull List<@NonNull Group> getGroupedRows() {
        if (this.groupedRows == null) {
            // only read group information on-demand
            List<@NonNull Group> newGroupedRows = new ArrayList<>();
            
            Deque<Integer> groupedRows = new ArrayDeque<>();
            int groupLevel = 0;
            int previousRow = -1;
            int lastNonEmptyRow = 0;
            
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
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
                        newGroupedRows.add(new Group(groupingStart, previousRow));
                        groupLevel--;
                    }
                }

                previousRow++;
                
                // check if row is empty
                boolean hasContent = false;
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (!hasContent && cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() != CellType.BLANK && !currentCell.toString().isEmpty()) {
                        hasContent = true;
                    }
                }
                if (hasContent) {
                    lastNonEmptyRow++;
                }
            }
            
            while (groupLevel > 0) {
                // Group ends at the last line
                Integer groupingStart = groupedRows.pollFirst();
                int lastRow = Math.min(previousRow, lastNonEmptyRow - 1);
                newGroupedRows.add(new Group(groupingStart, lastRow));
                groupLevel--;
            }
            
            this.groupedRows = newGroupedRows;
        }
        
        return notNull(Collections.unmodifiableList(this.groupedRows));
    }
    
    /**
     * Returns all grouped rows which are relevant for the specified index.
     * The elements are sorted in descending order of the start index, thus, the most inner group comes first,
     * the most outer group comes last. 
     * 
     * @param rowIndex A 0-based index for which the groups shall be returned.
     * @return A list of grouped rows, may be empty.
     */
    public @NonNull List<@NonNull Group> getRowGroups(int rowIndex) {
        List<@NonNull Group> relevantGroups = new ArrayList<>();
        for (Group rowGroup : getGroupedRows()) {
            if (rowGroup.getStartIndex() <= rowIndex && rowGroup.getEndIndex() >= rowIndex) {
                relevantGroups.add(rowGroup);
            }
        }
        
        // Sorts elements by start index in descending order
        relevantGroups.sort((g1, g2) -> Integer.compare(g2.getStartIndex(), g1.getStartIndex()));
        
        return notNull(Collections.unmodifiableList(relevantGroups));
    }
    
    @Override
    public void close() {
        // no need to close anything, just reset the iterator
        this.rowIterator = notNull(sheet.rowIterator());
        currentRowNumber = 0;
    }

    @Override
    public @NonNull String @Nullable [] readNextRow() throws IOException {
        @NonNull String[] result = null;
        
        List<String> rowContents;
        boolean isEmpty = true;
        boolean isEnd = false;

        // don't directly increment this.currentRowNumber
        // only set the value, if we actually find a non-empty line (this is important at the end of the file)
        int currentRowNumberCopy = this.currentRowNumber;
        
        do {
            rowContents = new ArrayList<>();
            
            if (!this.rowIterator.hasNext()) {
                isEnd = true; // to break the loop
                
            } else {
                Row currentRow = this.rowIterator.next();
                currentRowNumberCopy++;
                
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    
                    // Handle missing/undefined cells
                    while (currentCell.getColumnIndex() > rowContents.size()) {
                        rowContents.add("");
                    }
                    
                    String value;
                    switch (currentCell.getCellTypeEnum()) {
                    case STRING:
                        value = currentCell.getStringCellValue();
                        break;
                    case NUMERIC:
                        value = Double.toString(currentCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        value = Boolean.toString(currentCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        value = currentCell.getCellFormula();
                        break;
                    default: 
                        // getStringCellValue() returns "" for empty cells
                        value = currentCell.getStringCellValue();
                        break;
                    }
                    
                    isEmpty &= value == null;
                    rowContents.add(value);
                }
            }
            
        } while (!isEnd && (isEmpty && ignoreEmptyRows));
        
        if (!isEnd) {
            this.currentRowNumber = currentRowNumberCopy;
            
            // Handle missing/undefined cells at the end of row
            while (rowContents.size() < nColumns) {
                rowContents.add("");
            }
            
            result = rowContents.toArray(new @NonNull String[0]);
        }
        
        return result;
    }
    
    @Override
    public int getLineNumber() {
        return currentRowNumber;
    }

}
