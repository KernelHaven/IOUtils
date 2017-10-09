package net.ssehub.kernel_haven.io.excel;

import java.io.IOException;

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
        if (null != fields) {
            Row row = sheet.createRow(currentRow++);
            for (int i = 0; i < fields.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(fields[i]);
            }
        }
    }

}
