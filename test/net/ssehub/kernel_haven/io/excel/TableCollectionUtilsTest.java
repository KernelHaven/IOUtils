package net.ssehub.kernel_haven.io.excel;

import static org.junit.Assert.assertThat;

import java.io.File;
import java.io.IOException;

import org.hamcrest.CoreMatchers;
import org.junit.Test;

import net.ssehub.kernel_haven.util.io.ITableCollection;
import net.ssehub.kernel_haven.util.io.TableCollectionUtils;

/**
 * Tests the {@link TableCollectionUtils} (they should be able to handle Excel now that this plugin is available).
 * 
 * @author Adam
 */
public class TableCollectionUtilsTest {

    /**
     * Tests whether the {@link TableCollectionUtils} factory correctly creates Excel collections.
     * 
     * @throws IOException unwanted.
     */
    @Test
    public void testXls() throws IOException {
        ITableCollection collection = TableCollectionUtils.openFile(new File("test.xls"));
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
        ITableCollection collection = TableCollectionUtils.openFile(new File("test.xlsx"));
        assertThat(collection, CoreMatchers.instanceOf(ExcelBook.class));
        collection.close();
    }

}
