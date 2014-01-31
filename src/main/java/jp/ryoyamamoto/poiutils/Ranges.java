/*
 * Copyright 2014 Ryo Yamamoto
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 *     
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package jp.ryoyamamoto.poiutils;

import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

/**
 * Utility methods for {@link CellRangeAddress}.
 * 
 * @author Ryo Yamamoto
 */
public class Ranges {

    /*
     * Public methods
     */
    /**
     * Gets the reference of the upper-left most cell in the range.
     * 
     * @param range
     *            the range of cells
     * @return the reference of the upper-left most cell in the range
     */
    public static CellReference getFirstCellReference(CellRangeAddress range) {
        return toAreaReference(range).getFirstCell();
    }

    /**
     * Gets the reference of the lower-right most cell in the range.
     * 
     * @param range
     *            the range of cells
     * @return the reference of the lower-right most cell in the range
     */
    public static CellReference getLastCellReference(CellRangeAddress range) {
        return toAreaReference(range).getLastCell();
    }

    /**
     * Gets the references of all cells in the range.
     * 
     * @param range
     *            the range of cells
     * @return the references of all cells in the range
     */
    public static CellReference[] getCellReferences(CellRangeAddress range) {
        return toAreaReference(range).getAllReferencedCells();
    }

    /*
     * Private methods
     */
    /**
     * Converts a range to a {@link AreaReference}.
     * 
     * @param range
     *            the range of cells
     * @return the reference of the area
     */
    private static AreaReference toAreaReference(CellRangeAddress range) {
        return new AreaReference(range.formatAsString());
    }
}
