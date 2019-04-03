/*
 * Copyright 2017-2019 University of Hildesheim, Software Systems Engineering
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package net.ssehub.kernel_haven.io.excel;

import static org.junit.Assert.assertThat;

import java.io.File;
import java.io.IOException;

import org.hamcrest.CoreMatchers;
import org.junit.Test;

import net.ssehub.kernel_haven.util.io.ITableCollection;
import net.ssehub.kernel_haven.util.io.TableCollectionReaderFactory;

/**
 * Tests the {@link TableCollectionReaderFactory} (they should be able to handle Excel now that this plugin is
 * available).
 * 
 * @author Adam
 */
public class TableCollectionFactoryTest {

    /**
     * Tests whether the {@link TableCollectionReaderFactory} factory correctly creates Excel collections.
     * 
     * @throws IOException unwanted.
     */
    @Test
    public void testXls() throws IOException {
        ITableCollection collection = TableCollectionReaderFactory.INSTANCE.openFile(new File("test.xls"));
        assertThat(collection, CoreMatchers.instanceOf(ExcelBook.class));
        collection.close();
    }
    
    /**
     * Tests whether the {@link TableCollectionReaderFactory} factory correctly creates Excel collections.
     * 
     * @throws IOException unwanted.
     */
    @Test
    public void testXlsx() throws IOException {
        ITableCollection collection = TableCollectionReaderFactory.INSTANCE.openFile(new File("test.xlsx"));
        assertThat(collection, CoreMatchers.instanceOf(ExcelBook.class));
        collection.close();
    }

}
