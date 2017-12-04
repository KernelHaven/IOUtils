package net.ssehub.kernel_haven.io.excel;

import static org.hamcrest.CoreMatchers.hasItem;
import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.junit.Assert;
import org.junit.BeforeClass;
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
    private static final File TMPFOLDER = new File(AllTests.TESTDATA, "tmpFiles");

    @BeforeClass
    public static void setUpBeforeClass() {
        if (TMPFOLDER.exists()) {
            try {
                FileUtils.deleteDirectory(TMPFOLDER);
            } catch (IOException e) {
                Assert.fail("Could not clear temp directory: " + TMPFOLDER.getAbsolutePath());
            }
        }
        
        TMPFOLDER.mkdirs();
    }
    
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
     * Tests the correct retrieval of cell values, even if intermediate cells are undefined.
     * 
     */
    @Test
    public void testHandleUndefinedCells()  {
        String[][] allRows = loadSheet("UndefinedIntermediateCell.xlsx", "UndefinedCell");
        
        Assert.assertEquals("Expected 2 rows", 2, allRows.length);
        Assert.assertEquals("Expected 3 entries in 1st row", 3, allRows[0].length);
        Assert.assertEquals("Expected 3 entries in 2nd row", 3, allRows[1].length);
        Assert.assertEquals("Value 1", allRows[1][0]);
        Assert.assertEquals("", allRows[1][1]);
        Assert.assertEquals("Value 3", allRows[1][2]);
    }
    
    /**
     * Tests the correct retrieval of cell values, even if cells at the end of the row are undefined.
     */
    @Test
    public void testHandleUndefinedCellsAtEndOfRow()  {
        String[][] allRows = loadSheet("UndefinedLastCell.xlsx", "UndefinedCell");
        
        Assert.assertEquals("Expected 2 rows", 2, allRows.length);
        Assert.assertEquals("Expected 3 entries in 1st row", 3, allRows[0].length);
        Assert.assertEquals("Expected 3 entries in 2nd row", 3, allRows[1].length);
        Assert.assertEquals("Value 1", allRows[1][0]);
        Assert.assertEquals("Value 2", allRows[1][1]);
        Assert.assertEquals("", allRows[1][2]);
    }
    
    
    /**
     * Loads the specified sheet content from the specified workbook.
     * @param fileName The workbook to load, a path relative to {@link AllTests#TESTDATA}.
     * @param sheetName The name of the sheet to load.
     * @return Will return the content of the sheet (if it cannot be loaded, the test will fail already at this part).
     */
    private String[][] loadSheet(String fileName, String sheetName) {
        ExcelSheetReader reader = null;
        File inputFile = new File(AllTests.TESTDATA, fileName);
        String[][] content = null;
        try (ExcelBook book = new ExcelBook(inputFile, true)) {
            assertThat(book.getTableNames(), hasItem(sheetName));
            reader = book.getReader(sheetName);
            Assert.assertNotNull(reader);
            
            content = reader.readFull();
        } catch (IllegalStateException | IOException | FormatException e) {
            Assert.fail(e.getMessage());
        }
        
        Assert.assertNotNull(content);
        return content;
    }
    
    /**
     * 
     * @throws IOException if an error occurs while reading the data (Must not occur during testing)
     * @throws FormatException if the contents of the file cannot be parsed (Must not occur during testing)
     * @throws IllegalStateException If the workbook given is password protected (Must not occur during testing)
     */
    @Test
    public void testWriteSingleSheet() throws IllegalStateException, IOException, FormatException {
        File newWorkbook = new File(TMPFOLDER, "testWriteSingleSheet.xlsx");
        Assert.assertFalse(newWorkbook.exists());
        String sheetName = "newSheet";
        
        // Create empty book
        ExcelBook book = new ExcelBook(newWorkbook);
        Assert.assertEquals(0, book.getTableNames().size());
        
        // Create first sheet
        ExcelSheetWriter writer = book.getWriter(sheetName);
        Assert.assertEquals(1, book.getTableNames().size());
        Assert.assertTrue(book.getTableNames().contains(sheetName));
        writer.writeRow("A", "Test");
        
        // Write contents to file system
        book.close();
        
        // Test that correct content was written
        try (ExcelBook writtenBook = new ExcelBook(newWorkbook)) {
            ExcelSheetReader reader = writtenBook.getReader(sheetName);
            String[][] content = reader.readFull();
            Assert.assertEquals(1, content.length);
            String[] row1 = content[0];
            Assert.assertEquals(2, row1.length);
            Assert.assertEquals("A", row1[0]);
            Assert.assertEquals("Test", row1[1]);
            writtenBook.close();
        }
    }
    
    /**
     * 
     * @throws IOException if an error occurs while reading the data (Must not occur during testing)
     * @throws FormatException if the contents of the file cannot be parsed (Must not occur during testing)
     * @throws IllegalStateException If the workbook given is password protected (Must not occur during testing)
     */
    @Test
    public void testWriteMultipleSheets() throws IllegalStateException, IOException, FormatException {
        File newWorkbook = new File(TMPFOLDER, "testWriteMultipleSheets.xlsx");
        Assert.assertFalse(newWorkbook.exists());
        String sheetName1 = "newSheet1";
        String sheetName2 = "newSheet2";
        
        // Create empty book
        ExcelBook book = new ExcelBook(newWorkbook);
        Assert.assertEquals(0, book.getTableNames().size());
        
        // Create first sheet
        ExcelSheetWriter writer = book.getWriter(sheetName1);
        Assert.assertEquals(1, book.getTableNames().size());
        Assert.assertTrue(book.getTableNames().contains(sheetName1));
        Assert.assertFalse(book.getTableNames().contains(sheetName2));
        writer.writeRow("A", "Test");
        writer.close();
        
        // Create second sheet
        writer = book.getWriter(sheetName2);
        Assert.assertEquals(2, book.getTableNames().size());
        Assert.assertTrue(book.getTableNames().contains(sheetName1));
        Assert.assertTrue(book.getTableNames().contains(sheetName2));
        writer.writeRow("Another", "Test");
        writer.writeRow("With", "2 Rows");
        writer.close();
        
        // Write contents to file system
        book.close();
        
        // Test that correct content was written
        try (ExcelBook writtenBook = new ExcelBook(newWorkbook)) {
            ExcelSheetReader reader = writtenBook.getReader(sheetName1);
            String[][] content = reader.readFull();
            Assert.assertEquals(1, content.length);
            String[] row1 = content[0];
            Assert.assertEquals(2, row1.length);
            Assert.assertEquals("A", row1[0]);
            Assert.assertEquals("Test", row1[1]);
            
            reader = writtenBook.getReader(sheetName2);
            content = reader.readFull();
            Assert.assertEquals(2, content.length);
            row1 = content[0];
            Assert.assertEquals(2, row1.length);
            Assert.assertEquals("Another", row1[0]);
            Assert.assertEquals("Test", row1[1]);
            String[] row2 = content[1];
            Assert.assertEquals(2, row2.length);
            Assert.assertEquals("With", row2[0]);
            Assert.assertEquals("2 Rows", row2[1]);

            writtenBook.close();
        }
    }
    
    /**
     * Opens and then closes an empty excel book. This test was created for an out-of-range bug when closing an empty
     * book.
     * 
     * @throws IOException unwanted.
     * @throws FormatException unwanted.
     * @throws IllegalStateException unwanted.
     */
    @Test
    public void testOpenAndCloseNonExistingBook() throws IOException, IllegalStateException, FormatException {
        new ExcelBook(new File("testdata/DoesntExist.xlsx")).close();
    }
    
    /**
     * Tests the {@link ExcelBook#getFiles()} method.
     * 
     * @throws IOException unwanted.
     * @throws IllegalStateException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    public void testGetFiles() throws IOException, IllegalStateException, FormatException {
        try (ExcelBook book = new ExcelBook(new File("test.xls"))) {
            HashSet<File> expected = new HashSet<>();
            expected.add(new File("test.xls"));
            assertThat(book.getFiles(), is(expected));
        }
    }
    
    @Test(expected = UnsupportedOperationException.class)
    public void testWriteToExisting() throws IOException, IllegalStateException, FormatException {
        try (ExcelBook book = new ExcelBook(new File("testdata/Existing.xlsx"))) {
            book.getWriter("test").close();
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
