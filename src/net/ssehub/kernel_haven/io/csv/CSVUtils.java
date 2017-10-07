package net.ssehub.kernel_haven.io.csv;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;

import org.apache.commons.io.FileUtils;

import net.ssehub.kernel_haven.util.Logger;

/**
 * Utility functions for CSV file handling.
 * @author El-Sharkawy
 * 
 * @deprecated Use classes in main infrastructure instead.
 *
 */
@Deprecated
public class CSVUtils {
    
    /**
     * Avoids instantiation.
     */
    private CSVUtils() {}
    
    /**
     * Determines if the specified CSV-file has a separator specification and returns the specified character.
     * @param csvFile The file to read, should exist, otherwise <tt>null</tt> will be returned.
     * 
     * @return The specified separator or <tt>null</tt> if no separator was defined.
     */
    static Character determineSeparator(File csvFile) {
        Character separator = null;
        try {
            String line = FileUtils.readLines(csvFile, Charset.defaultCharset()).get(0);
            if (line != null && line.startsWith("sep")) {
                separator = line.charAt(line.length() - 1);
            }
        } catch (IOException e) {
            Logger.get().logWarning("Error occured while trying to determine delimiter of CSV file: " + e.getMessage());
        }
        
        return separator;
    }

}
