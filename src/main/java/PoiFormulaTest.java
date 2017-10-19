import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * <p>请注意：POI 中 在内存中修改了单元格数据，相应 <code>Formula 单元格数据</code> 还是之前的值，并不会动态改变。</p>
 * <p>推测：POI 在读取的时候就已经对 <code>Formula 单元格数据</code> 进行了计算，之后不会再次计算，除非手动刷新。</p>
 *
 * @author piumnl
 * @version 1.0.0
 * @since on 2017-10-19.
 */
public class PoiFormulaTest {

    public static void main(String[] args) {
        PoiFormulaTest fts = new PoiFormulaTest();
        System.out.println(fts.readFormulaValue());
    }

    public String readFormulaValue() {
        try (InputStream xlsxFile = this.getClass().getResourceAsStream("/formula.xlsx")){
            XSSFWorkbook hw = new XSSFWorkbook(xlsxFile);
            XSSFSheet hsheet = hw.getSheet("poi test");

            XSSFRow row = hsheet.getRow(2);
            XSSFCell cell = row.createCell(2);
            cell.setCellValue(1);
            row = hsheet.getRow(1);
            cell = row.createCell(2);
            cell.setCellValue(1);
            row = hsheet.getRow(0);
            cell = row.createCell(2);
            cell.setCellValue(1);

            // 此行代码用于手动控制计算所有 Formula单元格数据。
            // 不调用此方法则修改了单元格后不能实时获取 Formula单元格数据。
            XSSFFormulaEvaluator.evaluateAllFormulaCells(hw);

            XSSFRow hrow = hsheet.getRow(0);
            XSSFCell hcell = hrow.getCell(0);
            return this.getCellValue(hcell);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

    private String getCellValue(XSSFCell cell) {
        String value = null;
        if (cell != null) {
            switch (cell.getCellType()) {
                case XSSFCell.CELL_TYPE_FORMULA:
                    try {
                        value = String.valueOf(cell.getNumericCellValue());
                    } catch (IllegalStateException e) {
                        value = String.valueOf(cell.getRichStringCellValue());
                    }
                    break;
                case XSSFCell.CELL_TYPE_NUMERIC:
                    value = String.valueOf(cell.getNumericCellValue());
                    break;
                case XSSFCell.CELL_TYPE_STRING:
                    value = String.valueOf(cell.getRichStringCellValue());
                    break;
                default:
            }
        }

        return value;
    }
}
