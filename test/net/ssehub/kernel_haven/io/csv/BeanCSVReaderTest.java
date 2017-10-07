package net.ssehub.kernel_haven.io.csv;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;

import net.ssehub.kernel_haven.io.AllTests;

/**
 * Tests the {@link BeanCSVReader}.
 * @author El-Sharkawy
 *
 */
@Deprecated
public class BeanCSVReaderTest {

    /**
     * Tests that the delimiter specification of a file is used and the line is ignored while reading the contents.
     * @throws IOException
     */
    @Test
    public void testFileWithDelimiter() throws IOException {
        BeanCSVReader<VariableValueBean> reader = new BeanCSVReader<VariableValueBean>(new File(AllTests.TESTDATA, "CSV_with_Separator.csv"), VariableValueBean.class);
        List<VariableValueBean> contents = reader.readAll();
        
        Assert.assertNotNull("Contents of file wasn't read at all.", contents);
        Assert.assertEquals("Unexpected number of rows detected.", 2, contents.size());
        for (int i = 0; i < contents.size(); i++) {
            try {
                Integer.valueOf(contents.get(i).getValue());
            } catch (NumberFormatException exc) {
                Assert.fail("Value column wasn't read correctly: " + exc.getMessage());
            }
        }
    }
}
