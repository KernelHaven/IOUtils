package net.ssehub.kernel_haven.io.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

/**
 * Data representation for a sheet of an Excel document.
 * This sheet is designed to be read-only.
 * @author El-Sharkawy
 *
 */
public class ReadonlySheet implements Iterable<Object[]> {
    private String sheetName;
    private List<Object[]> contents = new ArrayList<>();
    private List<Group> groupedRows = new ArrayList<>();
    
    /**
     * Constructor for a named sheet.
     * @param sheetName The name of the sheet.
     */
    ReadonlySheet(String sheetName) {
        this.sheetName = sheetName;
    }
    
    /**
     * Adds a new row to the sheet.
     * @param rowContents The values of the row, should not be <tt>null</tt>.
     */
    void addRow(Object[] rowContents) {
        contents.add(rowContents);
    }
    
    /**
     * Adds a row grouping.
     * @param rowStart The first row of a group (0-based index).
     * @param rowEnd The last row of a group (0-based index).
     */
    public void addRowGrouping(int rowStart, int rowEnd) {
        groupedRows.add(new Group(rowStart, rowEnd));
    }
    
    /**
     * Returns an unmodifiable list of grouped rows of the document.
     * @return A list of grouped rows, may be empty.
     */
    public List<Group> getGroupedRows() {
        return Collections.unmodifiableList(groupedRows);
    }
    
    /**
     * Returns all grouped rows which are relevant for the specified index.
     * @param rowIndex A 0-based index for which the groups shall be returned.
     * @return A list of grouped rows, may be empty.
     */
    public List<Group> getRowGroups(int rowIndex) {
        List<Group> relevantGroups = new ArrayList<>();
        for (Group rowGroup : groupedRows) {
            if (rowGroup.getStartIndex() <= rowIndex && rowGroup.getEndIndex() >= rowIndex) {
                relevantGroups.add(rowGroup);
            }
        }
        
        return Collections.unmodifiableList(relevantGroups);
    }
    
    /**
     * Returns the number of stored rows.
     * @return A value &ge; 0.
     */
    public int getNumberOfRows() {
        return contents.size();
    }
    
    /**
     * Returns the specified row.
     * @param rowIndex A 0-based index. 
     * @return The contents of the specified row.
     */
    public Object[] getRow(int rowIndex) {
        return contents.get(rowIndex);
    }

    @Override
    public Iterator<Object[]> iterator() {
        Iterator<Object[]> delegate = contents.iterator();
        
        // Return a read only iterator
        return new Iterator<Object[]>() {
            @Override
            public boolean hasNext() {
                return delegate.hasNext();
            }

            @Override
            public Object[] next() {
                return delegate.next();
            }

        };
    }

    /**
     * Returns the name of the sheet.
     * @return The name of the sheet.
     */
    public String getSheetName() {
        return sheetName;
    }
}
