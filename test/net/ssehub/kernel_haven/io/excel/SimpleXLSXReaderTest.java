package net.ssehub.kernel_haven.io.excel;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;

import net.ssehub.kernel_haven.io.AllTests;
import net.ssehub.kernel_haven.util.FormatException;


public class SimpleXLSXReaderTest {
    
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
        Assert.assertEquals(1, firstGroup.getStartIndex());
        Assert.assertEquals(2, firstGroup.getEndIndex());
        Group secondGroup = groupedRows.get(1);
        Assert.assertEquals(3, secondGroup.getStartIndex());
        Assert.assertEquals(5, secondGroup.getEndIndex());
    }

}
