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

import java.util.Collections;
import java.util.List;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

/**
 * Utility methods for {@link Sheet}.
 * 
 * @author Ryo Yamamoto
 */
public class Sheets {

    /**
     * Merges two or more adjacent cells.
     * <p>
     * Only the data in the upper-left cell of a range will remain in the merged
     * cell. Data in other cells of the range will be deleted.
     * </p>
     * 
     * @param sheet
     *            the sheet that the range is on.
     * @param range
     *            the range to merge.
     */
    public static void merge(Sheet sheet, CellRangeAddress range) {
        boolean copied = false;
        Cell upperLeftCell = Sheets.getCell(sheet,
                Ranges.getFirstCellReference(range));
        for (CellReference reference : Ranges.getCellReferences(range)) {
            Cell cell = Sheets.getCell(sheet, reference);
            if (copied == false && Cells.isNotBlank(cell)) {
                Cells.copy(cell, upperLeftCell);
                copied = true;
            }
            if (cell != upperLeftCell) {
                Cells.clear(cell);
            }
        }
        sheet.addMergedRegion(range);
    }

    /**
     * Gets the hyperlinks on the sheet.
     * 
     * @param sheet
     *            the sheet.
     * @return the hyperlinks on the sheet.
     */
    @SuppressWarnings("unchecked")
    public static List<Hyperlink> getHyperlinks(Sheet sheet) {
        try {
            return (List<Hyperlink>) FieldUtils.readDeclaredField(sheet,
                    "hyperlinks", true);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    /**
     * Gets the cell at the specified reference.
     * 
     * @param sheet
     *            the sheet.
     * @param reference
     *            the reference of the cell.
     * @return the cell at the specified reference.
     */
    public static Cell getCell(Sheet sheet, CellReference reference) {
        return getCell(sheet, reference.getRow(), reference.getCol());
    }

    /**
     * Gets the cell at the specified row and column.
     * 
     * @param sheet
     *            the sheet.
     * @param row
     *            the row of the cell.
     * @param col
     *            the column of the cell.
     * @return the cell at the specified row and column.
     */
    public static Cell getCell(Sheet sheet, int row, int col) {
        return sheet.getRow(row).getCell(col);
    }
}
