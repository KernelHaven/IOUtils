package net.ssehub.kernel_haven.io.excel;

import static org.hamcrest.CoreMatchers.hasItem;
import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.junit.Assert.assertThat;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.BeforeClass;
import org.junit.Test;

import net.ssehub.kernel_haven.util.FormatException;
import net.ssehub.kernel_haven.util.null_checks.NonNull;

/**
 * Tests the {@link ExcelBook} and {@link ExcelSheetReader} classes.
 *
 * @author El-Sharkawy
 * @author Adam
 */
public class ExcelBookTest {
    
    private static final File TESTDATA = new File("testdata");
    
    private static final File TMPFOLDER = new File(TESTDATA, "tmpFiles");

    /**
     * Creates the {@link #TMPFOLDER}.
     */
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
        File inputFile = new File(TESTDATA, "GroupedValues.xlsx");
        
        try (ExcelBook book = new ExcelBook(inputFile, true)) {
            assertThat(book.getTableNames(), hasItem("Test Sheet"));
            
            ExcelSheetReader reader = book.getReader("Test Sheet");
            
            assertThat(reader.getLineNumber(), is(0));
            String[][] allRows = reader.readFull();
            assertThat(allRows.length, is(6));
            assertThat(reader.getLineNumber(), is(6));
            
            List<Group> groupedRows = reader.getGroupedRows();
            Assert.assertEquals(2, groupedRows.size());
            Group firstGroup = groupedRows.get(0);
            assertGroup(firstGroup, 1, 2);
            Group secondGroup = groupedRows.get(1);
            assertGroup(secondGroup, 4, 5);
            
            reader.close();
        }
    }
    
    /**
     * Tests the correct retrieval of groups for specified rows.
     * 
     * @throws IllegalStateException Should not occur, otherwise the tested Excel file is password protected.
     * @throws IOException Should not occur, otherwise the tested Excel document cannot be opened.
     * @throws FormatException Should not occur, otherwise the tested Excel cannot be parsed
     */
    @Test
    public void testGetRowGroups() throws IllegalStateException, IOException, FormatException {
        File inputFile = new File(TESTDATA, "GroupedValues.xlsx");
        
        try (ExcelBook book = new ExcelBook(inputFile, true)) {
            assertThat(book.getTableNames(), hasItem("Test Sheet"));
            
            ExcelSheetReader reader = book.getReader("Test Sheet");

            assertThat(reader.getLineNumber(), is(0));
            String[][] allRows = reader.readFull();
            assertThat(allRows.length, is(6));
            assertThat(reader.getLineNumber(), is(6));
            
            List<Group> groups = reader.getRowGroups(0);
            assertThat(groups.size(), is(0));
            
            groups = reader.getRowGroups(1);
            assertThat(groups.size(), is(1));
            assertGroup(groups.get(0), 1, 2);
            
            groups = reader.getRowGroups(2);
            assertThat(groups.size(), is(1));
            assertGroup(groups.get(0), 1, 2);
            
            groups = reader.getRowGroups(3);
            assertThat(groups.size(), is(0));
            
            groups = reader.getRowGroups(4);
            assertThat(groups.size(), is(1));
            assertGroup(groups.get(0), 4, 5);
            
            groups = reader.getRowGroups(5);
            assertThat(groups.size(), is(1));
            assertGroup(groups.get(0), 4, 5);
            
            groups = reader.getRowGroups(6);
            assertThat(groups.size(), is(0));
            
            reader.close();
        }
    }
    
    /**
     * Tests the correct retrieval of groups for specified rows. This sheet has nested groups
     * 
     * @throws IllegalStateException Should not occur, otherwise the tested Excel file is password protected.
     * @throws IOException Should not occur, otherwise the tested Excel document cannot be opened.
     * @throws FormatException Should not occur, otherwise the tested Excel cannot be parsed
     */
    @Test
    public void testGetRowGroupsNested() throws IllegalStateException, IOException, FormatException {
        File inputFile = new File(TESTDATA, "GroupedValues2.xlsx");
        
        try (ExcelBook book = new ExcelBook(inputFile, true)) {
            assertThat(book.getTableNames(), hasItem("Test Sheet"));
            
            ExcelSheetReader reader = book.getReader("Test Sheet");

            assertThat(reader.getLineNumber(), is(0));
            String[][] allRows = reader.readFull();
            assertThat(allRows.length, is(6));
            assertThat(reader.getLineNumber(), is(6));
            
            List<Group> groups = reader.getRowGroups(0);
            assertThat(groups.size(), is(0));
            
            groups = reader.getRowGroups(1);
            assertThat(groups.size(), is(2));
            assertGroup(groups.get(0), 1, 2);
            assertGroup(groups.get(1), 1, 5);
            
            groups = reader.getRowGroups(3);
            assertThat(groups.size(), is(1));
            assertGroup(groups.get(0), 1, 5);
            
            reader.close();
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
    
    /**
     * Tests the groups point to existing rows and won't cause {@link IndexOutOfBoundsException}s.
     * @throws IllegalStateException Should not occur, otherwise the tested Excel file is password protected.
     * @throws IOException Should not occur, otherwise the tested Excel document cannot be opened.
     * @throws FormatException Should not occur, otherwise the tested Excel cannot be parsed
     */
    @Test
    public void testGroupedRowsAccess() throws IllegalStateException, IOException, FormatException {
        File inputFile = new File(TESTDATA, "GroupedValues2.xlsx");
        
        try (ExcelBook book = new ExcelBook(inputFile, true)) {
            assertThat(book.getTableNames(), hasItem("Test Sheet"));
            
            ExcelSheetReader reader = book.getReader("Test Sheet");

            assertThat(reader.getLineNumber(), is(0));
            String[][] allRows = reader.readFull();
            assertThat(allRows.length, is(6));
            assertThat(reader.getLineNumber(), is(6));
            
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
            
            reader.close();
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
    private String[][] loadSheet(String fileName, @NonNull String sheetName) {
        ExcelSheetReader reader = null;
        File inputFile = new File(TESTDATA, fileName);
        String[][] content = null;
        try (ExcelBook book = new ExcelBook(inputFile, true)) {
            assertThat(book.getTableNames(), hasItem(sheetName));
            reader = book.getReader(sheetName);
            Assert.assertNotNull(reader);

            assertThat(reader.getLineNumber(), is(0));
            content = reader.readFull();
            assertThat(reader.getLineNumber(), is(content.length));
            reader.close();
        } catch (IllegalStateException | IOException | FormatException e) {
            Assert.fail(e.getMessage());
        }
        
        Assert.assertNotNull(content);
        return content;
    }
    
    /**
     * Test writing a single sheet.
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
        writer.close();
        
        // Write contents to file system
        book.close();
        
        // Test that correct content was written
        try (ExcelBook writtenBook = new ExcelBook(newWorkbook)) {
            ExcelSheetReader reader = writtenBook.getReader(sheetName);
            assertThat(reader.getLineNumber(), is(0));
            String[][] content = reader.readFull();
            assertThat(reader.getLineNumber(), is(1));
            Assert.assertEquals(1, content.length);
            String[] row1 = content[0];
            Assert.assertEquals(2, row1.length);
            Assert.assertEquals("A", row1[0]);
            Assert.assertEquals("Test", row1[1]);
            reader.close();
        }
    }
    
    /**
     * Tests writing multiple sheets.
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
            assertThat(reader.getLineNumber(), is(0));
            String[][] content = reader.readFull();
            Assert.assertEquals(1, content.length);
            assertThat(reader.getLineNumber(), is(1));
            String[] row1 = content[0];
            Assert.assertEquals(2, row1.length);
            Assert.assertEquals("A", row1[0]);
            Assert.assertEquals("Test", row1[1]);
            reader.close();
            
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
            reader.close();
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
    @SuppressWarnings("null")
    public void testGetFiles() throws IOException, IllegalStateException, FormatException {
        try (ExcelBook book = new ExcelBook(new File("testdata/Existing.xlsx"))) {
            HashSet<File> expected = new HashSet<>();
            expected.add(new File("testdata/Existing.xlsx"));
            assertThat(book.getFiles(), is(expected));
        }
    }
    
    /**
     * Tests that writing to an existing (and thus read-only) book throws an exception.
     * 
     * @throws IOException unwanted.
     * @throws IllegalStateException unwanted.
     * @throws FormatException unwanted.
     */
    @Test(expected = UnsupportedOperationException.class)
    public void testWriteToExisting() throws IOException, IllegalStateException, FormatException {
        try (ExcelBook book = new ExcelBook(new File("testdata/Existing.xlsx"))) {
            book.getWriter("test").close();
        }
    }
    
    /**
     * Tests that the getAllSheetReaders() method correctly returns all sheet readers.
     * 
     * @throws IOException unwanted.
     * @throws IllegalStateException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    @SuppressWarnings("null")
    public void testGetAllSheetReaders() throws IOException, IllegalStateException, FormatException {
        try (ExcelBook book = new ExcelBook(new File("testdata/MultipleSheets.xlsx"))) {
            List<ExcelSheetReader> readers = book.getAllSheetReaders();
            
            assertThat(readers.get(0).getSheetName(), is("Sheet1"));
            assertThat(readers.get(1).getSheetName(), is("Sheet2"));
            assertThat(readers.get(2).getSheetName(), is("Sheet3"));
            

            assertThat(readers.get(0).getLineNumber(), is(0));
            assertThat(readers.get(0).readFull(), is(new String[][] {{"Sheet", "One"}}));
            assertThat(readers.get(0).getLineNumber(), is(1));
            
            assertThat(readers.get(1).getLineNumber(), is(0));
            assertThat(readers.get(1).readFull(), is(new String[][] {{"Sheet", "Two"}}));
            assertThat(readers.get(1).getLineNumber(), is(1));
            
            assertThat(readers.get(2).getLineNumber(), is(0));
            assertThat(readers.get(2).readFull(), is(new String[][] {{"Sheet", "Three"}}));
            assertThat(readers.get(2).getLineNumber(), is(1));
            
            for (ExcelSheetReader reader : readers) {
                reader.close();
            }
        }
    }
    
    /**
     * Tests that reading an empty sheet works correctly. 
     * 
     * @throws IOException unwanted.
     * @throws IllegalStateException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    @SuppressWarnings("null")
    public void testReadEmptySheet() throws IOException, IllegalStateException, FormatException {
        try (ExcelBook book = new ExcelBook(new File("testdata/EmptySheet.xlsx"))) {
            try (ExcelSheetReader reader = book.getReader(0)) {
                assertThat(reader.getLineNumber(), is(0));
                assertThat(reader.readFull(), is(new String[0][]));
                assertThat(reader.getLineNumber(), is(0));
            }
        }
    }
    
    /**
     * Tests that reading an invalid file correctly throws an exception.
     * 
     * @throws IOException wanted.
     * @throws IllegalStateException unwanted.
     * @throws FormatException unwanted.
     */
    @Test(expected = IOException.class)
    public void testReadCorrupted() throws IOException, IllegalStateException, FormatException {
        ExcelBook book = new ExcelBook(new File("testdata/Corrupted.xls"));
        book.close();
    }
    
    /**
     * Tests that trying to create the same sheet (same name) twice is handled correctly.
     * 
     * @throws IOException unwanted.
     * @throws IllegalStateException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    @SuppressWarnings("null")
    public void testCreateSameSheetTwice() throws IOException, IllegalStateException, FormatException {
        File dst = new File("testdata/tmp.xls");
        try (ExcelBook book = new ExcelBook(dst)) {
            
            ExcelSheetWriter writer = book.getWriter("Sheet");
            writer.writeRow("Test", "Data");
            writer.close();
            
            writer = book.getWriter("Sheet");
            writer.writeRow("Other", "Test", "Data");
            writer.close();

            ExcelSheetReader reader = book.getReader("Sheet");
            assertThat(reader.getLineNumber(), is(0));
            assertThat(reader.readFull(), is(new String[][] {{"Other", "Test", "Data"}}));
            assertThat(reader.getLineNumber(), is(1));
            
        } finally {
            dst.delete();
        }
    }
    
    /**
     * Tests that the {@link ExcelSheetReader} can handled different content types.
     * 
     * @throws IOException unwanted.
     * @throws IllegalStateException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    @SuppressWarnings("null")
    public void testReadDifferentContentTypes() throws IOException, IllegalStateException, FormatException {
        try (ExcelBook book = new ExcelBook(new File("testdata/DifferentContentTypes.xlsx"))) {
            ExcelSheetReader reader = book.getReader(0);

            assertThat(reader.getLineNumber(), is(0));
            assertThat(reader.readFull(), is(new String[][] {
                {"String", "Text"},
                {"Numeric", "1.0"},
                {"Boolean", "true"},
                {"Formula", "3+2"},
                {"Blank", ""},
                {"Error", "4/0"},
            }));
            assertThat(reader.getLineNumber(), is(6));
            
            reader.close();
        }
    }
    
    /**
     * Tests that writing really long field names is handled correctly.
     * 
     * @throws IOException unwanted.
     * @throws IllegalStateException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    @SuppressWarnings("null")
    public void testWriteLongField() throws IOException, IllegalStateException, FormatException {
        File dst = new File("testdata/tmpLongFields.xls");
        final int length = SpreadsheetVersion.EXCEL2007.getMaxTextLength() + 200;
        
        try (ExcelBook book = new ExcelBook(dst)) {
            
            StringBuilder str = new StringBuilder();
            for (int i = 0; i < length; i++) {
                str.append('a');
            }
            
            ExcelSheetWriter writer = book.getWriter("Sheet");
            writer.writeObject(str.toString());
            writer.close();
            
            ExcelSheetReader reader = book.getReader("Sheet");
            assertThat(reader.getLineNumber(), is(0));
            String[] row = reader.readNextRow();
            assertThat(reader.getLineNumber(), is(1));
            assertThat(row.length, is(2));
            assertThat(row[0].length(), is(SpreadsheetVersion.EXCEL2007.getMaxTextLength()));
            assertThat(row[1].length(), is(200));

            assertThat(reader.readNextRow(), nullValue());
            assertThat(reader.getLineNumber(), is(1));
            reader.close();
            
        } finally {
            dst.delete();
        }
    }

    /**
     * Tests that writing different data types gets formatted correctly.
     * 
     * @throws IOException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    public void testWriteDifferentTypes() throws IOException, FormatException {
        File dst = new File("testdata/tmpDifferentTypes.xlsx");
        
        try (ExcelBook book = new ExcelBook(dst)) {
            
            ExcelSheetWriter writer = book.getWriter("Sheet");
            writer.writeRow("String", "A String Value");
            writer.writeRow("Integer", 13);
            writer.writeRow("Double", -13.5);
            writer.writeRow("Null", null);
            writer.writeRow("Boolean(s)", true, false);
            writer.close();
            
            ExcelSheetReader reader = book.getReader("Sheet");
            assertThat(reader.getLineNumber(), is(0));
            String[][] content = reader.readFull();
            assertThat(reader.getLineNumber(), is(5));
            
            assertThat(content, is(new String[][] {
                {"String", "A String Value"},
                {"Integer", "13.0"},
                {"Double", "-13.5"},
                {"Null", ""},
                {"Boolean(s)", "true", "false"},
            }));
            
            reader.close();
        } finally {
            dst.delete();
        }
        
    }
    
    /**
     * Tests writing a header line.
     * 
     * @throws IOException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    public void testWriteHeader() throws IOException, FormatException {
        File dst = new File("testdata/tmpWriteHeader.xlsx");
        
        try (ExcelBook book = new ExcelBook(dst)) {
            
            ExcelSheetWriter writer = book.getWriter("Sheet");
            writer.writeHeader("Context", "Value");
            writer.writeRow("A", "1");
            writer.writeRow("B", "2");
            writer.writeRow("C", "3");
            writer.close();
            
            ExcelSheetReader reader = book.getReader("Sheet");
            assertThat(reader.getLineNumber(), is(0));
            String[][] content = reader.readFull();
            assertThat(reader.getLineNumber(), is(4));
            
            assertThat(content, is(new String[][] {
                {"Context", "Value"},
                {"A", "1"},
                {"B", "2"},
                {"C", "3"},
            }));
            
            reader.close();
            
        } finally {
            dst.delete();
        }
        
    }
    
    /**
     * Tests writing a meta data (author, title, date).
     * 
     * @throws IOException unwanted.
     * @throws FormatException unwanted.
     */
    @Test
    public void testWriteMetadata() throws IOException, FormatException {
        File dst = new File("testdata/testWriteMetadata.xlsx");
        dst.deleteOnExit();
        String analysisName = "MetaAnalysis";

        dst.delete();
        Assert.assertFalse(dst.exists());
        try (ExcelBook book = new ExcelBook(dst)) {
            
            ExcelSheetWriter writer = book.getWriter(analysisName);
            writer.writeHeader("Context", "Value");
            writer.close();
            
        }
        
        try (XSSFWorkbook readMetadata = new XSSFWorkbook(new FileInputStream(dst))) {   
            POIXMLProperties props = readMetadata.getProperties();
            POIXMLProperties.CoreProperties coreProp = props.getCoreProperties();
            
            // Creator
            Assert.assertEquals("KernelHaven", coreProp.getCreator());
            
            // Title: Main analysis name (= 1st sheet) + date (in human readable form (not Adam readable form!))
            Date date = Calendar.getInstance().getTime();
            SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");
            String title = analysisName + " " + sdf.format(date);
            Assert.assertEquals(title, coreProp.getTitle());
        }
    }
    
}
