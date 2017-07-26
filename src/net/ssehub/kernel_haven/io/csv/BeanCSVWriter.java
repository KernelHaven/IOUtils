package net.ssehub.kernel_haven.io.csv;

import java.io.Closeable;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.List;

import com.opencsv.bean.StatefulBeanToCsv;
import com.opencsv.bean.StatefulBeanToCsvBuilder;
import com.opencsv.exceptions.CsvException;

/**
 * A bean-based CSV writer. Needs a bean, which is annotated with {@link com.opencsv.bean.CsvBindByName}
 * and/or {@link com.opencsv.bean.CsvBindByPosition} annotations to write the CSV file.<br/><br/>
 * 
 * <b><font color="red">Important: </font></b> Use the {@link #close()} method after writing the last entry.
 * @author El-Sharkawy
 *
 * @param <D> The bean, which is used as data type for writing the information.
 */
public class BeanCSVWriter<D> implements Closeable {
    
    private Writer fWriter;
    private StatefulBeanToCsv<D> writer;
    
    /**
     * Default constructor for this writer.
     * @param destFile The file to be written, will overwrite existing files.
     * @param dataType The bean type, must be the class instance of the used generic
     * @param separator The separator for the CSV file to be used, may be <tt>null</tt> than the platform default will
     *     be used.
     * 
     * @throws IOException If the file exists but is a directory rather than
     *                  a regular file, does not exist but cannot be created,
     *                  or cannot be opened for any other reason
     */
    @SuppressWarnings("unchecked")
    public BeanCSVWriter(File destFile, Class<D> dataType, Character separator) throws IOException {
        fWriter = new FileWriter(destFile);
        
        StatefulBeanToCsvBuilder<D> builder = new StatefulBeanToCsvBuilder<>(fWriter);
        if (null != separator) {
            builder = builder.withSeparator(separator);
        }
        
        writer = builder.build();
    }

    /**
     * Writes a single row/entry.
     * @param entry The data for one row to be written
     * 
     * @throws CsvException If a field of the bean is annotated improperly, an unsupported data type is supposed to be
     *     written, or a required field is <tt>null</tt>.
     */
    public void writeLine(D entry) throws CsvException {
        writer.write(entry);
    }
    
    /**
     * Writes all entries and closes the writer afterwards, using the {@link #close()} method.
     * @param entries The entries (rows) to be written 
     * 
     * @throws CsvException If a field of the bean is annotated improperly, an unsupported data type is supposed to be
     *     written, or a required field is <tt>null</tt>.
     * @throws IOException If an I/O error occurs
     */
    public void writeAll(List<D> entries) throws CsvException, IOException {
        writer.write(entries);
        close();
    }
    
    /**
     * Closes the writer, should be called after the last element was written.
     * @throws IOException If an I/O error occurs
     */
    public void close() throws IOException {
        fWriter.close();
        writer = null;
    }
}
