package net.ssehub.kernel_haven.io;

import java.io.File;
import java.io.IOException;

import net.ssehub.kernel_haven.io.excel.ExcelBook;
import net.ssehub.kernel_haven.util.FormatException;
import net.ssehub.kernel_haven.util.io.ITableCollection;
import net.ssehub.kernel_haven.util.io.csv.CsvFileSet;

/**
 * Utilities for {@link ITableCollection}s.
 *  
 * @author Adam
 */
public class TableCollectionUtils {

    /**
     * Don't allow instances.
     */
    private TableCollectionUtils() {
    }
    
    /**
     * Creates an {@link ITableCollection} for a given file, based on the file suffix:
     * <ul>
     *  <li><b>.csv</b>: A {@link CsvFileSet} with the single input file is created.</li>
     *  <li><b>.xlsx</b>: An {@link ExcelBook} is created.</li>
     * </ul>
     * Any other suffix is not supported.
     * 
     * @param file The file to create an {@link ITableCollection} for.
     * @return The {@link ITableCollection} for the given file.
     * 
     * @throws IOException If creating the {@link ITableCollection} fails, or the file suffix is not supported.
     */
    public static ITableCollection openExcelOrCsvCollection(File file) throws IOException {
        ITableCollection result;
        String fileName = file.getName();
        
        if (fileName.endsWith(".csv")) {
            result = new CsvFileSet(file);
            
        } else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls") ) {
            try {
                result = new ExcelBook(file);
            } catch (FormatException e) {
                throw new IOException(e);
            }
            
        } else {
            throw new IOException("Don't know how to read file " + file.getName());
        }
        
        return result;
    }
    
}
