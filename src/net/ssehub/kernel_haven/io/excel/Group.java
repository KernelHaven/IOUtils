package net.ssehub.kernel_haven.io.excel;

/**
 * Contains information about a grouped row or column.
 * @author El-Sharkawy
 *
 */
public class Group {
    private int startIndex;
    private int endIndex;
    
    /**
     * Sole constructor.
     * @param startIndex The first row/column of a group (0-based index).
     * @param endIndex The last row/column of a group (0-based index).
     */
    Group(int startIndex, int endIndex) {
        this.startIndex = startIndex;
        this.endIndex = endIndex;
    }

    /**
     * Returns the first row/column of the specified group (0-based index).
     * @return A value &ge; 0.
     */
    public int getStartIndex() {
        return startIndex;
    }
    
    /**
     * Returns the last row/column of the specified group (0-based index).
     * @return A value &ge; 0.
     */
    public int getEndIndex() {
        return endIndex;
    }
}
