package net.ssehub.kernel_haven.io.excel;

import static net.ssehub.kernel_haven.util.null_checks.NullHelpers.notNull;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import net.ssehub.kernel_haven.util.FormatException;
import net.ssehub.kernel_haven.util.Logger;
import net.ssehub.kernel_haven.util.io.ITableCollection;
import net.ssehub.kernel_haven.util.io.TableCollectionReaderFactory;
import net.ssehub.kernel_haven.util.io.TableCollectionWriterFactory;
import net.ssehub.kernel_haven.util.null_checks.NonNull;
import net.ssehub.kernel_haven.util.null_checks.Nullable;

/**
 * A wrapper around an excel book. A book contains several sheets. The individual sheets can be accessed through
 * {@link ExcelSheetReader}s, see {@link #getAllSheetReaders()} and {@link #getSheetReader(int)}.
 *
 * @author Adam
 * @author El-Sharkawy
 */
public class ExcelBook implements ITableCollection {
    
    static {
        // this static block is invoked by the infrastructure, see loadClasses.txt
        
        // register to TableCollectionReaderFactory
        TableCollectionReaderFactory.INSTANCE.registerHandler("xls", ExcelBook.class);
        TableCollectionReaderFactory.INSTANCE.registerHandler("xlsx", ExcelBook.class);
        
        // register to TableCollectionWriterFactory
        TableCollectionWriterFactory.INSTANCE.registerHandler("xls", ExcelBook.class);
        TableCollectionWriterFactory.INSTANCE.registerHandler("xlsx", ExcelBook.class);
    }
    
    private static final int ROW_WINDOW_SIZE = 10;
    
    /**
     * The read/write mode to open an {@link ExcelBook} with.
     */
    private static enum Mode {
        // Add further state if needed
        
        // Existing workbook, which shall not be changed
        READ_ONLY,
        
        // New workbook, read (temporary) data and add new data
        WRITE_NEW_WB;
    }
    
    private static final Logger LOGGER = Logger.get();
    
    private Workbook wb;
    private POIXMLProperties.CoreProperties wbProperties = null;
    
    private boolean ignoreEmptyRows;
    private Mode mode;
    private @NonNull File destinationFile;
    
    private Set<@NonNull ExcelSheetWriter> openWriters;
    
    /**
     * Default constructor for reading and writing a Excel documents (XLSX, XLS).
     * Will also consider empty lines during reading.
     * 
     * @param destinationFile An Excel document, which shall be parsed (if existing) or be written (if not existing).
     * 
     * @throws IOException if an error occurs while reading the data
     * @throws FormatException if the contents of the file cannot be parsed
     * @throws IllegalStateException If the workbook given is password protected
     */
    public ExcelBook(@NonNull File destinationFile) throws IOException, IllegalStateException, FormatException {
        this(destinationFile, false);
    }
    
    /**
     * Constructor for reading and writing a Excel documents (XLSX, XLS).
     * The second parameter may be used to specify whether empty lines shall be considered.
     * 
     * @param destinationFile An Excel document, which shall be parsed (if existing) or be written (if not existing).
     * @param ignoreEmptyRows <tt>true</tt> empty rows will be skipped, <tt>false</tt> all lines will be read.
     * 
     * @throws IOException if an error occurs while reading the data
     * @throws FormatException if the contents of the file cannot be parsed
     * @throws IllegalStateException If the workbook given is password protected
     */
    public ExcelBook(@NonNull File destinationFile, boolean ignoreEmptyRows) throws IOException, IllegalStateException,
        FormatException {
        
        this.ignoreEmptyRows = ignoreEmptyRows;
        this.destinationFile = destinationFile;
        openWriters = new HashSet<>();
        if (!destinationFile.exists()) {
            if (destinationFile.createNewFile()) {
                mode = Mode.WRITE_NEW_WB;
                SXSSFWorkbook wb = new SXSSFWorkbook(ROW_WINDOW_SIZE); 
                wb.setCompressTempFiles(true);
                this.wb = wb;
                
                // TODO: properties
//                POIXMLProperties xmlProps = ((XSSFWorkbook) wb).getProperties();    
//                wbProperties = xmlProps.getCoreProperties();
//                wbProperties.setCreator("KernelHaven");
            } else {
                throw new IOException("Specified file does not exist and could not be created: "
                    + destinationFile.getAbsolutePath());
            }
        } else {
            try {
                mode = Mode.READ_ONLY;
                /* Using a File object allows for lower memory consumption, while an InputStream requires more memory
                 * as it has to buffer the whole file.
                 */
                wb = WorkbookFactory.create(destinationFile, null, true);
            } catch (InvalidFormatException e) {
                throw new FormatException(e);
            }
        }
    }
    
    /**
     * Returns {@link ExcelSheetReader}s for all sheets in this book.
     * 
     * @return Readers for all sheets of the Excel document.
     */
    public synchronized @NonNull List<@NonNull ExcelSheetReader> getAllSheetReaders() {
        List<@NonNull ExcelSheetReader> result = new ArrayList<>();
        
        for (Sheet sheet : wb) {
            result.add(new ExcelSheetReader(notNull(sheet), ignoreEmptyRows));
        }
        
        return result;
    }
    
    /**
     * Returns a reader for the specified sheet.
     * 
     * @param index Index of the sheet number (0-based physical & logical)
     * @return A reader for the sheet at the provided index.
     * 
     * @throws IllegalArgumentException if the index is out of range (index
     *            &lt; 0 || index &gt;= getNumberOfSheets()).
     */
    public synchronized @NonNull ExcelSheetReader getReader(int index) {
        Sheet sheet = notNull(wb.getSheetAt(index));
        
        return new ExcelSheetReader(sheet, ignoreEmptyRows);
    }
    
    @Override
    public synchronized @NonNull Set<@NonNull String> getTableNames() throws IOException {
        Set<@NonNull String> result = new HashSet<>();
        
        for (Sheet sheet : wb) {
            result.add(notNull(sheet.getSheetName()));
        }
        
        return result;
    }
    
    @Override
    public synchronized @NonNull ExcelSheetReader getReader(@NonNull String name) throws IOException {
        ExcelSheetReader result = null;
        for (Sheet sheet : wb) {
            if (sheet.getSheetName().equals(name)) {
                result = new ExcelSheetReader(sheet, ignoreEmptyRows);
                break;
            }
        }
        
        if (result == null) {
            throw new IOException("Workbook does not contain a sheet with name " + name);
        }
        
        return result;
    }

    @Override
    public synchronized @NonNull ExcelSheetWriter getWriter(@NonNull String name) throws IOException {
        switch (mode) {
        case READ_ONLY:
            throw new UnsupportedOperationException("Sheet was oppened in read only mode: "
                + destinationFile.getAbsolutePath());
        case WRITE_NEW_WB:
            // falls through
        default:
            String safeName = WorkbookUtil.createSafeSheetName(name);
            IllegalArgumentException exception = null;
            Sheet sheet = null;
            try {
                sheet = wb.createSheet(safeName);
            } catch (IllegalArgumentException exc) {
                exception = exc;
                
                // either the sheet name is invalid, or a sheet with the same name exists already
                
                // check whether a sheet with the same name exists
                Sheet existing = wb.getSheet(safeName);
                if (existing != null) {
                    // if a sheet with the same name exists already, overwrite it (as specified in JavaDoc)
                    wb.removeSheetAt(wb.getSheetIndex(existing));
                    
                    // now try to create the sheet again
                    try {
                        sheet = wb.createSheet(safeName);
                    } catch (IllegalArgumentException exc2) {
                        exception = exc2;
                    }
                }
            }
            if (null == sheet) {
                String cause = null != exception ? ", cause: " + exception.getMessage() : "";
                throw new IOException("Could not create sheet \"" + safeName + "\"" + cause);
            }
            
            ExcelSheetWriter writer = new ExcelSheetWriter(this, sheet);
            openWriters.add(writer);
            return writer;
        }
    }
    
    @Override
    public synchronized void close() throws IOException {
        closingLoop();
        write();
        
        if (mode == Mode.WRITE_NEW_WB) {
            ((SXSSFWorkbook) wb).dispose();
        }
        
        wb.close();
    }

    @Override
    public @NonNull Set<@NonNull File> getFiles() throws IOException {
        Set<@NonNull File> result = new HashSet<>();
        result.add(destinationFile);
        return result;
    }
    
    /**
     * Returns the style for writing header elements in this complete workbook.
     * @return The same style instance for all sheets of the same workbook to highlight header elements,
     *     or <tt>null</tt> if this workbook was opened in read only mode.
     */
    synchronized @Nullable CellStyle getHeaderStyle() {
        CellStyle style = null;
        if (mode != Mode.READ_ONLY) {
            style = wb.createCellStyle();
            Font font = wb.createFont();
            font.setBold(true);
            style.setFont(font);
        }
        
        return style;
    }
    
    /**
     * Signals the the given writer is closed.
     * 
     * @param writer The writer which is closed and calls this method.
     * @throws IOException if the file exists but is a directory rather than a regular file, does not exist but cannot
     *     be created, or cannot be opened for any other reason, or if anything could not be written
     * @throws IllegalStateException If a future version of this class does not consider all possible states
     */
    synchronized void closeWriter(@NonNull ExcelSheetWriter writer) throws IOException, IllegalStateException {
        openWriters.remove(writer);
        // TODO: if we figure out whether we can flush the streaming workbook, do it here
    }

    /**
     * Writes the passed values as long as the document was not opened in read only mode.
     * 
     * @throws IOException if the file exists but is a directory rather than a regular file, does not exist but cannot
     *     be created, or cannot be opened for any other reason, or if anything could not be written
     * @throws IllegalStateException If a future version of this class does not consider all possible states
     */
    private void write() throws IOException, IllegalStateException {
        switch (mode) {
        case WRITE_NEW_WB:
            // check that there are sheets; if not, then no data was written and we do not create this book
            if (wb.getNumberOfSheets() > 0) {
                wb.setActiveSheet(0);
                BufferedOutputStream fileOut = new BufferedOutputStream(new FileOutputStream(destinationFile));
                wb.write(fileOut);
                
                if (null != wbProperties) {
                    String dateOfToday = null;
                    try {
                        Date date = Calendar.getInstance().getTime();
                        SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");
                        dateOfToday = sdf.format(date);
                    } catch (NumberFormatException | NullPointerException exc) {
                        LOGGER.logException("Could not determine date of today", exc);
                    }
                    // First sheet is usually named after the most relevant analysis
                    String title = (null != dateOfToday) ? wb.getSheetName(0) + " " + dateOfToday : wb.getSheetName(0);
                    wbProperties.setTitle(title);
                }
                
                fileOut.close();
            } else {
                // opening the workbook created an empty file; delete it, since we have no data to write
                destinationFile.delete();
            }
            
            // falls through
        case READ_ONLY:
            break;
        default:
            // Should not happen, this is only to ensure that future versions consider all states.
            throw new IllegalStateException("Unexpected close opperation for state: " + mode.name());
        }
    }
    
    /**
     * Will wait for open writers 5 seconds until it will close all writers.
     * Will also suppress but log all exceptions to avoid crashing of whole Workbook.
     */
    private void closingLoop() {
        // Wait for open writers, maybe they still receive data.
        int attemptNo = 0;
        while (!openWriters.isEmpty() && attemptNo < 5) {
            try {
                Thread.sleep(1000);
            } catch (InterruptedException e) {
                LOGGER.logWarning("Error while ExcelBook is waiting for its sheets: " + e.getMessage());
            }
            attemptNo++;
        }
        
        // Close open writers
        List<@NonNull ExcelSheetWriter> tmp = new ArrayList<>(openWriters);
        for (ExcelSheetWriter excelSheetWriter : tmp) {
            try {
                closeWriter(excelSheetWriter);
            } catch (IOException e) {
                LOGGER.logError("Error while writing sheet: " + e.getMessage());
            }
        }
    }
}
