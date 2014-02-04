package jp.ryoyamamoto.poiutils;

import static org.assertj.core.api.Assertions.assertThat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.junit.experimental.runners.Enclosed;
import org.junit.runner.RunWith;

@RunWith(Enclosed.class)
public class SheetsTest {

    public static class WhenYouMergeTwoOrMoreCells {

        private Sheet sheet;
        private CellRangeAddress range;

        @Before
        public void setUp() throws Exception {
            sheet = createWorkbook().getSheetAt(0);
            range = new CellRangeAddress(0, 1, 0, 1);
        }

        @Test
        public void onlyTheValueInTheUpperLeftCellShouldRemain() {
            Sheets.merge(sheet, range);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("A1");
            assertThat(sheet.getRow(0).getCell(1).getCellType()).isEqualTo(
                    Cell.CELL_TYPE_BLANK);
            assertThat(sheet.getRow(1).getCell(0).getCellType()).isEqualTo(
                    Cell.CELL_TYPE_BLANK);
            assertThat(sheet.getRow(1).getCell(1).getCellType()).isEqualTo(
                    Cell.CELL_TYPE_BLANK);
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet();
            Row row1 = sheet.createRow(0);
            row1.createCell(0).setCellValue("A1");
            row1.createCell(1).setCellValue("B1");
            Row row2 = sheet.createRow(1);
            row2.createCell(0).setCellValue("A2");
            row2.createCell(1).setCellValue("B2");
            return workbook;
        }
    }
}
