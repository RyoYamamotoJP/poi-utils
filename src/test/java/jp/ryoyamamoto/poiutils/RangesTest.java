package jp.ryoyamamoto.poiutils;

import static jp.ryoyamamoto.poiutils.Ranges.getCellReferences;
import static jp.ryoyamamoto.poiutils.Ranges.getFirstCellReference;
import static jp.ryoyamamoto.poiutils.Ranges.getLastCellReference;
import static jp.ryoyamamoto.poiutils.Ranges.toAreaReference;
import static org.assertj.core.api.Assertions.assertThat;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.junit.Before;
import org.junit.Test;
import org.junit.experimental.runners.Enclosed;
import org.junit.runner.RunWith;

@RunWith(Enclosed.class)
public class RangesTest {

    public static class WhenTheRangeIsA1B2 {

        private CellRangeAddress a1b2;
        private CellReference a1;
        private CellReference a2;
        private CellReference b1;
        private CellReference b2;

        @Before
        public void setUp() throws Exception {
            a1b2 = CellRangeAddress.valueOf("A1:B2");
            a1 = new CellReference("A1");
            a2 = new CellReference("A2");
            b1 = new CellReference("B1");
            b2 = new CellReference("B2");
        }

        @Test
        public void getFirstCellReferenceShouldReturnA1() {
            assertThat(getFirstCellReference(a1b2)).isEqualTo(a1);
        }

        @Test
        public void getLastCellReferenceShouldReturnB2() {
            assertThat(getLastCellReference(a1b2)).isEqualTo(b2);
        }

        @Test
        public void getReferencesShouldReturnTheReferencesOfTheCellsBetweenA1AndB2() {
            CellReference[] references = getCellReferences(a1b2);
            assertThat(references).contains(a1, a2, b1, b2);
        }

        @Test
        public void toAreaReferenceShouldReturnA1B2() {
            String areaReference = toAreaReference(a1b2).formatAsString();
            assertThat(areaReference).isEqualTo("A1:B2");
        }
    }
}
