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

/**
 * Contains information about a grouped row or column.
 * @author El-Sharkawy
 *
 */
public class Group {
    private int startIndex;
    private int endIndex;
    
    /**
     * Sole constructor.
     * @param startIndex The first row/column of a group (0-based index).
     * @param endIndex The last row/column of a group (0-based index).
     */
    Group(int startIndex, int endIndex) {
        this.startIndex = startIndex;
        this.endIndex = endIndex;
    }

    /**
     * Returns the first row/column of the specified group (0-based index).
     * @return A value &ge; 0.
     */
    public int getStartIndex() {
        return startIndex;
    }
    
    /**
     * Returns the last row/column of the specified group (0-based index).
     * @return A value &ge; 0.
     */
    public int getEndIndex() {
        return endIndex;
    }
    
    @Override
    public String toString() {
        // For Debugging only
        return "[" + startIndex + ";" + endIndex + "]";
    }
}
