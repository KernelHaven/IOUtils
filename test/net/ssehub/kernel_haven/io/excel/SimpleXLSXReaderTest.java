package net.ssehub.kernel_haven.io.excel;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;

import net.ssehub.kernel_haven.io.AllTests;
import net.ssehub.kernel_haven.util.FormatException;

/**
 * Tests the {@link SimpleXLSXReaderTest}.
 * @author El-Sharkawy
 *
 */
public class SimpleXLSXReaderTest {
    
    /**
     * Tests the correct retrieval of grouped rows.
     * @throws IllegalStateException Should not occur, otherwise the tested Excel file is password protected.
     * @throws IOException Should not occur, otherwise the tested Excel document cannot be opened.
     * @throws FormatException Should not occur, otherwise the tested Excel cannot be parsed
     */
    @Test
    public void testGroupedRows() throws IllegalStateException, IOException, FormatException {
        File inputFile = new File(AllTests.TESTDATA, "GroupedValues.xlsx");
        SimpleExcelReader reader = new SimpleExcelReader(inputFile, true);
        List<ReadonlySheet> sheets = reader.readAll();
        reader.close();
        
        Assert.assertEquals(1, sheets.size());
        ReadonlySheet sheet = sheets.get(0);
        Assert.assertEquals("Test Sheet", sheet.getSheetName());
        
        int rows = 0;
        Iterator<Object[]> itr = sheet.iterator();
        while (itr.hasNext()) {
            rows++;
            itr.next();
        }
        Assert.assertEquals(6, rows);
        
        List<Group> groupedRows = sheet.getGroupedRows();
        Assert.assertEquals(2, groupedRows.size());
        Group firstGroup = groupedRows.get(0);
        assertGroup(firstGroup, 1, 2);
        Group secondGroup = groupedRows.get(1);
        assertGroup(secondGroup, 4, 5);
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
