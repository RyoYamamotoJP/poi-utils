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

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Utility methods for {@link Cell}.
 * 
 * @author Ryo Yamamoto
 */
public class Cells {

    /*
     * Public methods
     */
    /**
     * Copies a cell, including formulas and their resulting values, comments,
     * and cell formats.
     * 
     * @param source
     *            the source cell.
     * @param target
     *            the target cell.
     */
    public static void copy(Cell source, Cell target) {
        if (source == null) {
            target = null;
            return;
        }

        copyCellValue(source, target);
        copyCellStyle(source, target);
        copyCellComment(source, target);
        copyHyperlink(source, target);
    }

    /**
     * Clears a cell to remove the cell contents (formulas and data), formats
     * (including number formats, conditional formats, and borders), and any
     * attached comments.
     * <p>
     * The cleared cell remain as a blank or an unformatted cell.
     * </p>
     * 
     * @param cell
     *            the cell to be cleared.
     */
    public static void clear(Cell cell) {
        if (cell == null) {
            return;
        }

        cell.setCellType(Cell.CELL_TYPE_BLANK);
    }

    /**
     * Returns {@code true} if the cell is blank or {@code null}.
     * 
     * @param cell
     *            the cell to test.
     * @return {@code true} if the cell is blank or {@code null}.
     */
    public static boolean isBlank(Cell cell) {
        return cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK;
    }

    /**
     * Returns {@code true} if the cell is not blank.
     * 
     * @param cell
     *            the cell to test.
     * @return {@code true} if the cell is not blank.
     */
    public static boolean isNotBlank(Cell cell) {
        return !isBlank(cell);
    }

    /*
     * Private methods
     */
    private static void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
        case Cell.CELL_TYPE_NUMERIC:
            target.setCellValue(source.getNumericCellValue());
            break;
        case Cell.CELL_TYPE_STRING:
            target.setCellValue(source.getRichStringCellValue());
            break;
        case Cell.CELL_TYPE_FORMULA:
            target.setCellFormula(source.getCellFormula());
            break;
        case Cell.CELL_TYPE_BLANK:
            break;
        case Cell.CELL_TYPE_BOOLEAN:
            target.setCellValue(source.getBooleanCellValue());
            break;
        case Cell.CELL_TYPE_ERROR:
            target.setCellErrorValue(source.getErrorCellValue());
            break;
        }
    }

    private static void copyCellStyle(Cell source, Cell target) {
        target.setCellStyle(source.getCellStyle());
    }

    private static void copyCellComment(Cell source, Cell target) {
        target.setCellComment(source.getCellComment());
    }

    private static void copyHyperlink(Cell source, Cell target) {
        if (source.getHyperlink() == null) {
            removeHyperlink(target);
        } else {
            target.setHyperlink(source.getHyperlink());
        }
    }

    private static void removeHyperlink(Cell cell) {
        if (cell == null || cell.getHyperlink() == null) {
            return;
        }

        Sheet sheet = getSheet(cell);
        List<Hyperlink> hyperlinks = Sheets.getHyperlinks(sheet);
        hyperlinks.remove(cell.getHyperlink());
    }

    private static Sheet getSheet(Cell cell) {
        return cell.getRow().getSheet();
    }
}
