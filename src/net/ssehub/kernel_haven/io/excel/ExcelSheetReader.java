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
    
    private @NonNull List<@NonNull String[]> contents;

    private @NonNull List<@NonNull Group> groupedRows;
    
    private boolean ignoreEmptyRows;
    
    /**
     * The current position for {@link #readNextRow()}. Reset when {@link #close()} is called.
     */
    private @NonNull Iterator<@NonNull String[]> iterator;
    
    private int currentRow;
    
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
        this.contents = new ArrayList<>();
        this.groupedRows = new ArrayList<>();
        
        read();
        
        this.iterator = notNull(contents.iterator()); 
    }
    
    /**
     * Reads the full sheet. After this, {@link #contents} is filled.
     */
    private void read() {
        // Retrieves only the number of entries in first column (unsure if this is detailed enough)
        int nColumns = 0;
        if (sheet.getRow(0) != null) {
            nColumns = sheet.getRow(0).getLastCellNum();
        }
        Iterator<Row> rowIterator = sheet.rowIterator();
        Deque<Integer> groupedRows = new ArrayDeque<>();
        int groupLevel = 0;
        int previousRow = -1;
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
                    this.groupedRows.add(new Group(groupingStart, previousRow));
                    groupLevel--;
                }
            }
            
            readSingleRow(nColumns, currentRow);
            
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

    /**
     * Reads a single row. Called inside {@link #read()}. Adds the result to {@link #contents}.
     * 
     * @param nColumns The expected number of columns (used to handle missing cells at the end of row).
     * @param currentRow The row to read.
     */
    private void readSingleRow(int nColumns, @NonNull Row currentRow) {
        List<String> rowContents = new ArrayList<>();
        Iterator<Cell> cellIterator = currentRow.iterator();
        boolean isEmpty = true;
        while (cellIterator.hasNext()) {
            Cell currentCell = cellIterator.next();
            
            // Handle missing/undefined cells
            while (currentCell.getColumnIndex() > rowContents.size()) {
                rowContents.add("");
            }
            
            String value = null;
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
                value = currentCell.getStringCellValue();
                break;
            }
            
            isEmpty &= value == null;
            rowContents.add(value);
        }
        
        if (!ignoreEmptyRows || !isEmpty) {
            // Handle missing/undefined cells at the end of row
            while (rowContents.size() < nColumns) {
                rowContents.add("");
            }
            this.contents.add(rowContents.toArray(new @NonNull String[0]));
        }
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
        return groupedRows;
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
        for (Group rowGroup : groupedRows) {
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
        iterator = notNull(contents.iterator());
        currentRow = 0;
    }

    @Override
    public @NonNull String @Nullable [] readNextRow() throws IOException {
        @NonNull String[] result;
        if (iterator.hasNext()) {
            result = iterator.next();
            currentRow++;
        } else {
            result = null;
        }
        return result;
    }
    
    @Override
    public int getLineNumber() {
        return currentRow;
    }

}
