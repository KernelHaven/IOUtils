package net.ssehub.kernel_haven.io;

import static org.junit.Assert.assertThat;

import java.io.File;
import java.io.IOException;

import org.hamcrest.CoreMatchers;
import org.junit.Test;

import net.ssehub.kernel_haven.io.excel.ExcelBook;
import net.ssehub.kernel_haven.util.io.ITableCollection;
import net.ssehub.kernel_haven.util.io.csv.CsvFileSet;

public class TableCollectionUtilsTest {

    /**
     * Tests whether the {@link TableCollectionUtils} factory correctly creates CSV collections.
     * 
     * @throws IOException unwanted.
     */
    @Test
    public void testCsv() throws IOException {
        ITableCollection collection = TableCollectionUtils.openExcelOrCsvCollection(new File("test.csv"));
        assertThat(collection, CoreMatchers.instanceOf(CsvFileSet.class));
        collection.close();
    }
    
    /**
     * Tests whether the {@link TableCollectionUtils} factory correctly creates Excel collections.
     * 
     * @throws IOException unwanted.
     */
    @Test
    public void testXls() throws IOException {
        ITableCollection collection = TableCollectionUtils.openExcelOrCsvCollection(new File("test.xls"));
        assertThat(collection, CoreMatchers.instanceOf(ExcelBook.class));
        collection.close();
    }
    
    /**
     * Tests whether the {@link TableCollectionUtils} factory correctly creates Excel collections.
     * 
     * @throws IOException unwanted.
     */
    @Test
    public void testXlsx() throws IOException {
        ITableCollection collection = TableCollectionUtils.openExcelOrCsvCollection(new File("test.xlsx"));
        assertThat(collection, CoreMatchers.instanceOf(ExcelBook.class));
        collection.close();
    }
    
    /**
     * Tests whether the {@link TableCollectionUtils} factory correctly throws an exception if an invalid file suffix
     * is passed to it.
     * 
     * @throws IOException wanted.
     */
    @Test(expected = IOException.class)
    public void testInvalidTsv() throws IOException {
        TableCollectionUtils.openExcelOrCsvCollection(new File("test.tsv"));
    }
    
    /**
     * Tests whether the {@link TableCollectionUtils} factory correctly throws an exception if an invalid file suffix
     * is passed to it.
     * 
     * @throws IOException wanted.
     */
    @Test(expected = IOException.class)
    public void testInvalidTxt() throws IOException {
        TableCollectionUtils.openExcelOrCsvCollection(new File("test.txt"));
    }
    
}
