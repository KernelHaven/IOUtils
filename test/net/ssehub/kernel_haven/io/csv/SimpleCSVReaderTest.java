package net.ssehub.kernel_haven.io.csv;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;

import net.ssehub.kernel_haven.io.AllTests;

/**
 * Tests the {@link SimpleCSVReader}.
 * @author El-Sharkawy
 *
 */
public class SimpleCSVReaderTest {

    /**
     * Tests that the delimiter specification of a file is used and the line is ignored while reading the contents.
     * @throws IOException
     */
    @Test
    public void testFileWithDelimiter() throws IOException {
        SimpleCSVReader reader = new SimpleCSVReader(new File(AllTests.TESTDATA, "CSV_with_Separator.csv"));
        List<String[]> contents = reader.readAll();
        reader.close();
        
        Assert.assertNotNull("Contents of file wasn't read at all.", contents);
        Assert.assertEquals("Unexpected number of rows detected.", 3, contents.size());
        
        for (int i = 0; i < contents.size(); i++) {
            Assert.assertEquals("Unexpected number of colums detected in row " + (i + 1), 2, contents.get(i).length);
        }
    }

}
