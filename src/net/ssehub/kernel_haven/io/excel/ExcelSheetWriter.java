package net.ssehub.kernel_haven.io.excel;

import static net.ssehub.kernel_haven.util.null_checks.NullHelpers.notNull;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.ssehub.kernel_haven.util.io.AbstractTableWriter;
import net.ssehub.kernel_haven.util.null_checks.NonNull;
import net.ssehub.kernel_haven.util.null_checks.Nullable;

/**
 * Writes a table to an existing sheet of an existing Excel workbook (XLS/XLSX-File).
 * @see <a href="https://poi.apache.org/spreadsheet/quick-guide.html">
 * https://poi.apache.org/spreadsheet/quick-guide.html</a>
 * @author El-Sharkawy
 *
 */
public class ExcelSheetWriter extends AbstractTableWriter {
    
    private static final int MAX_TEXT_LENGTH = SpreadsheetVersion.EXCEL2007.getMaxTextLength();
    
    private @NonNull Sheet sheet;
    private int currentRow;
    private @NonNull ExcelBook wb;
    
    ExcelSheetWriter(@NonNull ExcelBook wb, @NonNull Sheet sheet) {
        this.wb = wb;
        this.sheet = sheet;
        currentRow = sheet.getPhysicalNumberOfRows();
    }

    @Override
    public void close() throws IOException {
        /*
         * In principle no needed, closing operation is handled in Workbook.
         * However, flushing current data is possible
         */
        wb.flush(this);
    }

    @Override
    public void writeRow(@Nullable Object /*@NonNull*/ ... columns) throws IOException {
        // TODO: commented out @NonNull annotation because checkstyle can't parse it
        // make sure we don't modify the content while the workbook is writing to disk
        synchronized (wb) {
            List<CellValue> cellValues = prepareFields(columns);
            Row row = sheet.createRow(currentRow++);
            for (int i = 0; i < cellValues.size(); i++) {
                Cell cell = notNull(row.createCell(i));
                cellValues.get(i).applyTo(cell);
            }
        }
    }
    
    @Override
    public void writeHeader(@Nullable Object /*@NonNull*/ ... fields) throws IOException {
        // TODO: commented out @NonNull annotation because checkstyle can't parse it
        // make sure we don't modify the content while the workbook is writing to disk
        synchronized (wb) {
            List<CellValue> cellValues = prepareFields(fields);
            Row row = sheet.createRow(currentRow++);
            for (int i = 0; i < cellValues.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(wb.getHeaderStyle());
                cellValues.get(i).applyTo(cell);
            }
            sheet.createFreezePane(0, 1);
        }
    }
    
    /**
     * A single cell value, with a type, to be written into the sheet.
     */
    private static class CellValue {
        
        private @NonNull CellType type;
        
        private @Nullable Object value;
        
        public CellValue(@NonNull CellType type, @Nullable Object value) {
            this.type = type;
            this.value = value;
        }
        
        /**
         * Applies this value to the given cell. Properly sets the cell type.
         * 
         * @param cell The cell to apply this value to.
         * 
         * @throws ClassCastException If the type does not match the value. Shouldn't happen.
         */
        public void applyTo(@NonNull Cell cell) {
            cell.setCellType(type);
            
            switch (type) {
            case BLANK:
                // no need to set a value
                break;
                // only BLANK has null as possible value, so its safe to assume that value != null now.
                
            case NUMERIC:
                cell.setCellValue(((Number) notNull(value)).doubleValue());
                break;
                
            case BOOLEAN:
                cell.setCellValue((Boolean) notNull(value));
                break;
                
            default:
                cell.setCellValue(notNull(value).toString());
                break;
            }
        }
        
    }
    
    /**
     * Converts the given fields into {@link CellValue}s.
     * <br />
     * Splits text values, which are too long into separate fields to avoid {@link IllegalArgumentException}s.
     * Tries to split values at white space characters.
     * <a href="https://stackoverflow.com/a/31937583">https://stackoverflow.com/a/31937583</a>
     * 
     * @param fields The field values of a row to store.
     * @return The values to write, should be the same values unless there were some values to long.
     * 
     */
    private @NonNull List<CellValue> prepareFields(@Nullable Object /*@NonNull*/ ... fields) {
        // TODO: commented out @NonNull annotation because checkstyle can't parse it
        List<CellValue> result = new ArrayList<>();
        
        for (Object field : fields) {
            
            if (field == null) {
                // empty fields are "blank" type
                result.add(new CellValue(CellType.BLANK, null));
                
            } else if (field instanceof Number) {
                // numbers get the "numeric" type
                result.add(new CellValue(CellType.NUMERIC, field));
                
            } else if (field instanceof Boolean) {
                // booleans are "boolean" type
                result.add(new CellValue(CellType.BOOLEAN, field));
                
            } else {
                // everything else is a "string" type
                // strings may be too long, and thus need to be split up
                
                String fieldValue = field.toString();
                while (fieldValue.length() > MAX_TEXT_LENGTH) {
                    String firstPart = fieldValue.substring(0, MAX_TEXT_LENGTH);
                    
                    // Try to split at a white space
                    int pos = firstPart.lastIndexOf(' ');
                    if (pos == -1) {
                        pos = MAX_TEXT_LENGTH;
                    }
                    
                    firstPart = fieldValue.substring(0, pos);
                    result.add(new CellValue(CellType.STRING, firstPart));
                    pos = Math.min(pos, fieldValue.length() - 1);
                    fieldValue = fieldValue.substring(pos);
                }
                result.add(new CellValue(CellType.STRING, fieldValue));
            }
        }
            
        
        return result;
    }

}
