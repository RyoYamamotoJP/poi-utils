package jp.ryoyamamoto.poiutils;

import static org.assertj.core.api.Assertions.assertThat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.junit.experimental.runners.Enclosed;
import org.junit.runner.RunWith;

@RunWith(Enclosed.class)
public class SSTest {

    public static class WhenYouCopyFromANumericCellToAnother {

        private Cell source;
        private Cell target;

        @Before
        public void setUp() throws Exception {
            Sheet sheet = createWorkbook().getSheetAt(0);
            source = sheet.getRow(0).getCell(0);
            target = sheet.createRow(1).createCell(0);
        }

        @Test
        public void theCellTypeOfTheTargetCellShouldBeNumeric() {
            SS.copy(source, target);
            assertThat(target.getCellType()).isEqualTo(Cell.CELL_TYPE_NUMERIC);
        }

        @Test
        public void theNumericValueOfTheTargetCellShouldBeReadable() {
            SS.copy(source, target);
            assertThat(target.getNumericCellValue()).isEqualTo(1);
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            workbook.createSheet().createRow(0).createCell(0).setCellValue(1);
            return workbook;
        }
    }

    public static class WhenYouCopyFromAStringCellToAnother {

        private Cell source;
        private Cell target;

        @Before
        public void setUp() throws Exception {
            Sheet sheet = createWorkbook().getSheetAt(0);
            source = sheet.getRow(0).getCell(0);
            target = sheet.createRow(1).createCell(0);
        }

        @Test
        public void theCellTypeOfTheTargetCellShouldBeString() {
            SS.copy(source, target);
            assertThat(target.getCellType()).isEqualTo(Cell.CELL_TYPE_STRING);
        }

        @Test
        public void theStringValueOfTheTargetCellShouldBeReadable() {
            SS.copy(source, target);
            assertThat(target.getStringCellValue()).isEqualTo("string");
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            workbook.createSheet().createRow(0).createCell(0)
                    .setCellValue("string");
            return workbook;
        }
    }

    public static class WhenYouCopyFromAFormulaCellToAnother {

        private Cell source;
        private Cell target;

        @Before
        public void setUp() throws Exception {
            Sheet sheet = createWorkbook().getSheetAt(0);
            source = sheet.getRow(0).getCell(0);
            target = sheet.createRow(1).createCell(0);
        }

        @Test
        public void theCellTypeOfTheTargetCellShouldBeFormula() {
            SS.copy(source, target);
            assertThat(target.getCellType()).isEqualTo(Cell.CELL_TYPE_FORMULA);
        }

        @Test
        public void theFormulaOfTheTargetCellShouldBeReadable() {
            SS.copy(source, target);
            assertThat(target.getCellFormula()).isEqualTo("SUM(B1:C1)");
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            workbook.createSheet().createRow(0).createCell(0)
                    .setCellFormula("SUM(B1:C1)");
            return workbook;
        }
    }

    public static class WhenYouCopyFromABlankCellToAnother {

        private Workbook workbook;
        private Cell source;
        private Cell target;

        @Before
        public void setUp() throws Exception {
            workbook = createWorkbook();
            source = workbook.getSheetAt(0).getRow(0).getCell(0);
            target = workbook.getSheetAt(0).createRow(1).createCell(0);
        }

        @Test
        public void theCellTypeOfTheTargetCellShouldBeBlank() {
            SS.copy(source, target);
            assertThat(target.getCellType()).isEqualTo(Cell.CELL_TYPE_BLANK);
        }

        @Test
        public void theNumericValueOfTheTargetCellShouldBeZero() {
            SS.copy(source, target);
            assertThat(target.getNumericCellValue()).isZero();
        }

        @Test
        public void theStringValueOfTheTargetCellShouldBeEmpty() {
            SS.copy(source, target);
            assertThat(target.getStringCellValue()).isEqualTo("");
        }

        @Test
        public void theBooleanValueOfTheTargetCellShouldBeFalse() {
            SS.copy(source, target);
            assertThat(target.getBooleanCellValue()).isFalse();
        }

        @Test
        public void theStyleOfTheTargetCellShouldBeDefault() {
            SS.copy(source, target);
            assertThat(target.getCellStyle()).isEqualTo(
                    workbook.getCellStyleAt((short) 0));
        }

        @Test
        public void theCommentOfTheTargetCellShouldBeNull() {
            SS.copy(source, target);
            assertThat(target.getCellComment()).isNull();
        }

        @Test
        public void theHyperlinkOfTheTargetCellShouldBeNull() {
            SS.copy(source, target);
            assertThat(target.getHyperlink()).isNull();
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet();
            // source
            sheet.createRow(0).createCell(0);
            // target
            CreationHelper factory = workbook.getCreationHelper();
            ClientAnchor anchor = factory.createClientAnchor();
            Cell cell = sheet.createRow(1).createCell(0);
            cell.setCellValue(1);
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(CellStyle.ALIGN_CENTER);
            cell.setCellStyle(style);
            Drawing drawing = sheet.createDrawingPatriarch();
            Comment comment = drawing.createCellComment(anchor);
            comment.setString(factory.createRichTextString("Comment"));
            cell.setCellComment(comment);
            Hyperlink link = factory.createHyperlink(Hyperlink.LINK_URL);
            link.setAddress("https://github.com/RyoYamamotoJP");
            cell.setHyperlink(link);
            return workbook;
        }
    }

    public static class WhenYouCopyFromABooleanCellToAnother {

        private Cell source;
        private Cell target;

        @Before
        public void setUp() throws Exception {
            Sheet sheet = createWorkbook().getSheetAt(0);
            source = sheet.getRow(0).getCell(0);
            target = sheet.createRow(1).createCell(0);
        }

        @Test
        public void theCellTypeOfTheTargetCellShouldBeBoolean() {
            SS.copy(source, target);
            assertThat(target.getCellType()).isEqualTo(Cell.CELL_TYPE_BOOLEAN);
        }

        @Test
        public void theBooleanValueOfTheTargetCellShouldBeReadable() {
            SS.copy(source, target);
            assertThat(target.getBooleanCellValue()).isTrue();
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            workbook.createSheet().createRow(0).createCell(0)
                    .setCellValue(true);
            return workbook;
        }
    }

    public static class WhenYouCopyFromAErrorCellToAnother {

        private Cell source;
        private Cell target;

        @Before
        public void setUp() throws Exception {
            Sheet sheet = createWorkbook().getSheetAt(0);
            source = sheet.getRow(0).getCell(0);
            target = sheet.createRow(1).createCell(0);
        }

        @Test
        public void theCellTypeOfTheTargetCellShouldBeError() {
            SS.copy(source, target);
            assertThat(target.getCellType()).isEqualTo(Cell.CELL_TYPE_ERROR);
        }

        @Test
        public void theErrorValueOfTheTargetCellShouldBeReadable() {
            SS.copy(source, target);
            assertThat(target.getErrorCellValue()).isEqualTo(
                    FormulaError.DIV0.getCode());
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            workbook.createSheet().createRow(0).createCell(0)
                    .setCellErrorValue(FormulaError.DIV0.getCode());
            return workbook;
        }
    }

    public static class WhenYouCopyFromACellToAnother {

        private Workbook workbook;
        private Cell source;
        private Cell target;

        @Before
        public void setUp() throws Exception {
            workbook = createWorkbook();
            source = workbook.getSheetAt(0).getRow(0).getCell(0);
            target = workbook.getSheetAt(0).createRow(1).createCell(0);
        }

        @Test
        public void theStyleOfTheTargetCellShouldNotBeDefault() {
            SS.copy(source, target);
            assertThat(target.getCellStyle()).isNotEqualTo(
                    workbook.getCellStyleAt((short) 0));
        }

        @Test
        public void theCommentOfTheTargetCellShouldBeReadable() {
            SS.copy(source, target);
            assertThat(target.getCellComment().getString().getString())
                    .isEqualTo("Comment");
        }

        @Test
        public void theHyperlinkOfTheTargetCellShouldBeReadable() {
            SS.copy(source, target);
            assertThat(target.getHyperlink().getAddress()).isEqualTo(
                    "https://github.com/RyoYamamotoJP");
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            CreationHelper factory = workbook.getCreationHelper();
            ClientAnchor anchor = factory.createClientAnchor();
            Sheet sheet = workbook.createSheet();
            Cell cell = sheet.createRow(0).createCell(0);
            cell.setCellValue(1);
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(CellStyle.ALIGN_CENTER);
            cell.setCellStyle(style);
            Drawing drawing = sheet.createDrawingPatriarch();
            Comment comment = drawing.createCellComment(anchor);
            comment.setString(factory.createRichTextString("Comment"));
            cell.setCellComment(comment);
            Hyperlink link = factory.createHyperlink(Hyperlink.LINK_URL);
            link.setAddress("https://github.com/RyoYamamotoJP");
            cell.setHyperlink(link);
            return workbook;
        }
    }

    public static class WhenYouClearACell {

        private Cell target;

        @Before
        public void setUp() throws Exception {
            Sheet sheet = createWorkbook().getSheetAt(0);
            target = sheet.getRow(0).createCell(0);
        }

        @Test
        public void theCellTypeOfTheTargetCellShouldBeBlank() {
            SS.clear(target);
            assertThat(target.getCellType()).isEqualTo(Cell.CELL_TYPE_BLANK);
        }

        private Workbook createWorkbook() {
            Workbook workbook = new XSSFWorkbook();
            workbook.createSheet().createRow(0).createCell(0).setCellValue(1);
            return workbook;
        }
    }
}
