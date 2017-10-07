package net.ssehub.kernel_haven.io.excel;

import static org.hamcrest.CoreMatchers.hasItem;
import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;

import net.ssehub.kernel_haven.io.AllTests;
import net.ssehub.kernel_haven.util.FormatException;

/**
 * Tests the {@link ExcelBook} and {@link ExcelSheetReader} classes.
 *
 * @author El-Sharkawy
 * @author Adam
 */
public class ExcelBookTest {

    /**
     * Tests the correct retrieval of grouped rows.
     * 
     * @throws IllegalStateException Should not occur, otherwise the tested Excel file is password protected.
     * @throws IOException Should not occur, otherwise the tested Excel document cannot be opened.
     * @throws FormatException Should not occur, otherwise the tested Excel cannot be parsed
     */
    @Test
    public void testGroupedRows() throws IllegalStateException, IOException, FormatException {
        File inputFile = new File(AllTests.TESTDATA, "GroupedValues.xlsx");
        
        try (ExcelBook book = new ExcelBook(inputFile, true)) {
            assertThat(book.getTableNames(), hasItem("Test Sheet"));
            
            ExcelSheetReader reader = book.getReader("Test Sheet");
            
            String[][] allRows = reader.readFull();
            assertThat(allRows.length, is(6));
            
            List<Group> groupedRows = reader.getGroupedRows();
            Assert.assertEquals(2, groupedRows.size());
            Group firstGroup = groupedRows.get(0);
            assertGroup(firstGroup, 1, 2);
            Group secondGroup = groupedRows.get(1);
            assertGroup(secondGroup, 4, 5);
        }
    }
    
    /**
     * Tests the groups point to existing rows and won't cause {@link IndexOutOfBoundsException}s.
     * @throws IllegalStateException Should not occur, otherwise the tested Excel file is password protected.
     * @throws IOException Should not occur, otherwise the tested Excel document cannot be opened.
     * @throws FormatException Should not occur, otherwise the tested Excel cannot be parsed
     */
    @Test
    public void testGroupedRowsAccess() throws IllegalStateException, IOException, FormatException {
        File inputFile = new File(AllTests.TESTDATA, "GroupedValues2.xlsx");
        
        try (ExcelBook book = new ExcelBook(inputFile, true)) {
            assertThat(book.getTableNames(), hasItem("Test Sheet"));
            
            ExcelSheetReader reader = book.getReader("Test Sheet");
            
            String[][] allRows = reader.readFull();
            assertThat(allRows.length, is(6));
            
            List<Group> groupedRows = reader.getGroupedRows();
            Assert.assertEquals(3, groupedRows.size());
            for (Group group : groupedRows) {
                String[] row = null;
                
                int firstIndex = group.getStartIndex();
                try {
                    row = allRows[firstIndex];
                    Assert.assertNotNull("Illegal index for first row index: " + firstIndex, row);
                } catch (IndexOutOfBoundsException exc) {
                    Assert.fail("Group covers non existing rows, row " + firstIndex + " is probably smaller than 0: "
                        + exc.getMessage());
                }
                
                int lastIndex = group.getEndIndex();
                try {
                    row = allRows[lastIndex];
                    Assert.assertNotNull("Illegal index for last row index: " + lastIndex, row);
                } catch (IndexOutOfBoundsException exc) {
                    Assert.fail("Group covers non existing rows, row " + lastIndex + " is greater than "
                        + allRows.length + ": " + exc.getMessage());
                }
            }
        }
    }

    /**
     * Asserts the correct setting of the tested group.
     * @param group The group to test.
     * @param startIndex The expected first row of the group (starts a 0).
     * @param endIndex The expected last row of the group (starts a 0).
     */
    private void assertGroup(Group group, int startIndex, int endIndex) {
        Assert.assertEquals("Start index for Group " + group + " not as expected.", startIndex, group.getStartIndex());
        Assert.assertEquals("End index for Group " + group + " not as expected.", endIndex, group.getEndIndex());
    }
    
}
