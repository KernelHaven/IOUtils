package net.ssehub.kernel_haven.io.csv;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.util.List;

import com.opencsv.bean.CsvToBean;
import com.opencsv.bean.CsvToBeanBuilder;

/**
 * Uses a bean to read a CSV file. The bean must be annotated with {@link com.opencsv.bean.CsvBindByName}
 * and/or {@link com.opencsv.bean.CsvBindByPosition} annotations.
 * @param <R> The bean, which shall be used as return type.
 * 
 * @author El-Sharkawy
 *
 */
public class BeanCSVReader<R> {
    
    private CsvToBean<R> reader;
    
    /**
     * Default constructor of this reader.
     * @param inputFile The file to read.
     * @param beanClass The annotated bean, which specifies the structure of the CSV-file to be parsed.
     * 
     * @throws FileNotFoundException if the file does not exist, is a directory rather than a regular file,
     *     or for some other reason cannot be opened for reading.
     */
    @SuppressWarnings({ "unchecked", "rawtypes" })
    public BeanCSVReader(File inputFile, Class<R> beanClass) throws FileNotFoundException {
        // Checks if file exists at all
        FileReader fReader = new FileReader(inputFile);
        
        Character separator = CSVUtils.determineSeparator(inputFile);
        
        if (null != separator) {
            // Skip first line and use specified separator
            reader = new CsvToBeanBuilder(fReader).withType(beanClass).withSeparator(separator).withSkipLines(1)
                .build();
        } else {
            reader = new CsvToBeanBuilder(fReader).withType(beanClass).build();
        }
    }
    
    /**
     * Parses the input based on parameters already set through other methods.
     * @return A list of populated beans based on the input
     * @throws IllegalStateException In case of errors.
     */
    public List<R> readAll() {
        return reader.parse();
    }

}
