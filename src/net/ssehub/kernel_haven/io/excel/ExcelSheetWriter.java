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
    private ExcelBook wb;
    
    ExcelSheetWriter(Sheet sheet) {
        this.sheet = sheet;
        currentRow = sheet.getPhysicalNumberOfRows();
    }
    
    ExcelSheetWriter(ExcelBook wb, Sheet sheet) {
        this(sheet);
        this.wb = wb;
    }

    @Override
    public void close() throws IOException {
        /*
         * In principle no needed, closing operation is handled in Workbook.
         * However, flushing current data is possible
         */
        if (null != wb) {
            wb.flush(this);
        }
    }

    @Override
    public void writeRow(Object... columns) throws IOException {
        // make sure we don't modify the content while the workbook is writing to disk
        synchronized (wb) {
            List<String> cellValues = prepareFields(columns);
            if (null != cellValues) {
                Row row = sheet.createRow(currentRow++);
                for (int i = 0; i < cellValues.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(cellValues.get(i));
                }
            }
        }
    }
    
    @Override
    public void writeHeader(Object... fields) throws IOException {
        // make sure we don't modify the content while the workbook is writing to disk
        synchronized (wb) {
            List<String> cellValues = prepareFields(fields);
            if (null != cellValues) {
                Row row = sheet.createRow(currentRow++);
                for (int i = 0; i < cellValues.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellStyle(wb.getHeaderStyle());
                    cell.setCellValue(cellValues.get(i));
                }
                sheet.createFreezePane(0, 1);
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
    private List<String> prepareFields(Object... fields) {
        List<String> result = null;
        if (null != fields) {
            result = new ArrayList<>();
            
            for (int i = 0; i < fields.length; i++) {
                String fieldValue = fields[i] != null ? fields[i].toString() : "";
                while (fieldValue.length() > MAX_TEXT_LENGTH) {
                    String firstPart = fieldValue.substring(0, MAX_TEXT_LENGTH);
                    
                    // Try to split at a white space
                    int pos = firstPart.lastIndexOf(' ');
                    if (pos == -1) {
                        pos = MAX_TEXT_LENGTH;
                    }
                    
                    firstPart = fieldValue.substring(0, pos);
                    result.add(firstPart);
                    pos = Math.min(pos, fieldValue.length() - 1);
                    fieldValue = fieldValue.substring(pos);
                }
                result.add(fieldValue);
            }
            
        }
        
        return result;
    }

}
