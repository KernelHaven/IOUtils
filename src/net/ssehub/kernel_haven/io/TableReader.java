package net.ssehub.kernel_haven.io;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import net.ssehub.kernel_haven.io.csv.SimpleCSVReader;
import net.ssehub.kernel_haven.io.excel.ReadonlySheet;
import net.ssehub.kernel_haven.io.excel.SimpleExcelReader;
import net.ssehub.kernel_haven.util.FormatException;

/**
 * Reads from Excel (XLSX, XLS) and CSV files.
 * @author El-Sharkawy
 *
 */
public class TableReader {
    
    /**
     * Supported file types.
     * @author El-Sharkawy
     *
     */
    private static enum FILE_TYPE {
        EXCEL, CSV;
    }
    
    /**
     * {@link ReadonlySheet} for the CSV reader.
     * @author El-Sharkawy
     *
     */
    private static class Sheet extends ReadonlySheet {

        private Sheet(String sheetName, Collection<String[]> contents) {
            super(sheetName);
            for (String[] row : contents) {
                super.addRow(row);
            }
        }
        
    }
    
    private File file;
    private FILE_TYPE type;
    private Character csvDelimiter;
    
    /**
     * Default constructor for this class.
     * @param file The file to be read (either an Excel file or a CSV file).
     * @throws IOException If the specified file does not exist.
     * @throws IllegalArgumentException If the specified file is neither an an Excel file or a CSV file.
     */
    public TableReader(File file) throws IOException, IllegalArgumentException {
        this.file = file;
        if (!this.file.exists()) {
            throw new IOException(file.getAbsolutePath() + " does not exist");
        }
        String fileName = file.getName();
        int pos = fileName.lastIndexOf('.');
        if (-1 == pos) {
            throw new IllegalArgumentException("Could not determine file extension of " + fileName);
        }
        String extension = fileName.substring(pos + 1).toLowerCase();
        if ("csv".equals(extension)) {
            type = FILE_TYPE.CSV;
        } else if ("xls".equals(extension) || "xlsx".equals(extension)) {
            type = FILE_TYPE.EXCEL;
        } else {
            throw new IllegalArgumentException("File \"" + fileName + "\" is neither a CSV nor an Excel file.");
        }
        csvDelimiter = null;
    }
    
    /**
     * Constructor to specify a delimiter/separator for entries in a CSV file.
     * @param file The file to be read (either an Excel file or a CSV file).
     * @throws IOException If the specified file does not exist.
     * @throws IllegalArgumentException If the specified file is neither an an Excel file or a CSV file.
     */
    public TableReader(File file, char csvDelimiter) throws IOException, IllegalArgumentException {
        this(file);
        this.csvDelimiter = csvDelimiter;
    }

    /**
     * Reads and converts the given file into an {@link ReadonlySheet}.
     * @return The read data of the given file.
     * @throws IllegalStateException If the workbook given is password protected or if the file type is not supported
     *     by this reader (should not happen).
     * @throws IOException If an error occurs while reading the file contents.
     * @throws FormatException if the contents of the file cannot be parsed.
     */
    public List<ReadonlySheet> read() throws IllegalStateException, IOException, FormatException {
        List<ReadonlySheet> result;
        switch (type) {
        case EXCEL:
            SimpleExcelReader excelReader = new SimpleExcelReader(file);
            result = excelReader.readAll();
            break;
        case CSV:
            SimpleCSVReader csvReader = (null == csvDelimiter) ? new SimpleCSVReader(file)
                : new SimpleCSVReader(file, csvDelimiter);
            List<String[]> content = csvReader.readAll();
            result = new ArrayList<>();
            result.add(new Sheet(file.getName(), content));
            break;
        default:
            throw new IllegalStateException("Unsupported file type: " + file.getName());
        }
        
        return result;
    }
}
