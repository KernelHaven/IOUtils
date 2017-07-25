package net.ssehub.kernel_haven.io.csv;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.List;

import org.apache.commons.io.FileUtils;

import com.opencsv.CSVReader;
import com.opencsv.ICSVParser;

import net.ssehub.kernel_haven.util.Logger;

/**
 * A simplistic CSV reader, based on <a href="http://opencsv.sourceforge.net/">http://opencsv.sourceforge.net/</a>.
 * @author El-Sharkawy
 *
 */
public class SimpleCSVReader {
    
    private CSVReader reader;
    
    /**
     * Creates a new CSV reader for the given input file with default separator.
     * If first line contains a delimiter specification, this specification will be used instead of the default one.
     * @param inputFile The file to read.
     * 
     * @throws FileNotFoundException if the file does not exist, is a directory rather than a regular file,
     *     or for some other reason cannot be opened for reading.
     */
    public SimpleCSVReader(File inputFile) throws FileNotFoundException {
        // Checks if file exists at all
        FileReader fReader = new FileReader(inputFile);
        
        Character separator = null;
        try {
            String line = FileUtils.readLines(inputFile, Charset.defaultCharset()).get(0);
            if (line != null && line.startsWith("sep")) {
                separator = line.charAt(line.length() - 1);
            }
        } catch (IOException e) {
            Logger.get().logWarning("Error occured while trying to determine delimiter of CSV file: " + e.getMessage());
        }
        
        if (null != separator) {
            // Skip first line and use specified separator
            reader = new CSVReader(fReader, separator, ICSVParser.DEFAULT_QUOTE_CHARACTER, 1);
        } else {
            reader = new CSVReader(fReader);
        }
    }
    
    /**
     * Creates a new CSV reader for the given input file with default separator.
     * If first line contains a delimiter specification, this specification will be ignored.
     * @param inputFile The file to read.
     * @param delimiter The delimiter to use to separate distinct values/cells from each other. This will overwrite
     * a <tt>sep=</tt> specification.
     * 
     * @throws FileNotFoundException if the file does not exist, is a directory rather than a regular file,
     *     or for some other reason cannot be opened for reading.
     */
    public SimpleCSVReader(File inputFile, char delimiter) throws FileNotFoundException {
        // Checks if file exists at all
        FileReader fReader = new FileReader(inputFile);
        
        Character separator = null;
        try {
            String line = FileUtils.readLines(inputFile, Charset.defaultCharset()).get(0);
            if (line != null && line.startsWith("sep")) {
                separator = line.charAt(line.length() - 1);
            }
        } catch (IOException e) {
            Logger.get().logWarning("Error occured while trying to determine delimiter of CSV file: " + e.getMessage());
        }
        
        if (null != separator) {
            // Skip first line, but use specified delimiter
            reader = new CSVReader(fReader, delimiter, ICSVParser.DEFAULT_QUOTE_CHARACTER, 1);
        } else {
            reader = new CSVReader(fReader);
        }
    }
    
    /**
     * Reads the next line from the buffer and converts to a string array.
     * @return An array, containing the values of the current line. Won't be <tt>null</tt> unless the the end of the
     * file is reached.
     * 
     * @throws IOException If bad things happen during the read.
     */
    public String[] readLine() throws IOException {
        return reader.readNext();
    }

    /**
     * Reads the entire file into a List with each element being a String[] of
     * tokens.
     *
     * @return A List of String[], with each String[] representing a line of the
     * file.
     * @throws IOException If bad things happen during the read
     */
    public List<String[]> readAll() throws IOException {
        return reader.readAll();
    }

}
