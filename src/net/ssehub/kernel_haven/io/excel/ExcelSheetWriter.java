package net.ssehub.kernel_haven.io.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.ssehub.kernel_haven.util.io.AbstractTableWriter;

/**
 * Writes a table to an existing sheet of an existing Excel workbook (XLS/XLSX-File).
 * @see <a href="https://poi.apache.org/spreadsheet/quick-guide.html">
 * https://poi.apache.org/spreadsheet/quick-guide.html</a>
 * @author El-Sharkawy
 *
 */
public class ExcelSheetWriter extends AbstractTableWriter {
    
    private static final int MAX_TEXT_LENGTH = SpreadsheetVersion.EXCEL2007.getMaxTextLength();
    
    private Sheet sheet;
    private int currentRow;
    
    ExcelSheetWriter(Sheet sheet) {
        this.sheet = sheet;
        currentRow = sheet.getPhysicalNumberOfRows();
    }

    @Override
    public void close() throws IOException {
        // Not needed, closing operation is handled in Workbook.
    }

    @Override
    public void writeRow(String... fields) throws IOException {
        List<String> cellValues = prepareFields(fields);
        if (null != cellValues) {
            Row row = sheet.createRow(currentRow++);
            for (int i = 0; i < cellValues.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(cellValues.get(i));
            }
        }
    }
    
    /**
     * Splits text values, which are too long into separate fields to avoid {@link IllegalArgumentException}s.
     * Tries to split values at white space characters.
     * @param fields The field values of a row to store.
     * @return The values to write, should be the same values unless there were some values to long.
     * <a href="https://stackoverflow.com/a/31937583">https://stackoverflow.com/a/31937583</a>
     */
    private List<String> prepareFields(String... fields) {
        List<String> result = null;
        if (null != fields) {
            result = new ArrayList<>();
            
            for (int i = 0; i < fields.length; i++) {
                String fieldValue = fields[i];
                while (fieldValue.length() > MAX_TEXT_LENGTH) {
                    String firstPart = fieldValue.substring(0, MAX_TEXT_LENGTH);
                    
                    // Try to split at a white space
                    int pos = firstPart.lastIndexOf(' ');
                    if (pos == -1) {
                        pos = MAX_TEXT_LENGTH;
                    }
                    
                    firstPart = fieldValue.substring(0, pos);
                    result.add(firstPart);
                    fieldValue = fieldValue.substring(pos);
                }
                result.add(fieldValue);
            }
            
        }
        
        return result;
    }

}
