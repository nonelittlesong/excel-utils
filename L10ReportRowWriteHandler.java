import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.constant.OrderConstant;
import com.alibaba.excel.util.BooleanUtils;
import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.handler.context.RowWriteHandlerContext;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class L10ReportRowWriteHandler implements RowWriteHandler {
    private final XSSFCellStyle[] styles = new XSSFCellStyle[30];

    @Override
    public int order() {
        return OrderConstant.FILL_STYLE + 1;
    }

    @Override
    public void afterRowDispose(RowWriteHandlerContext context) {
        Row row = context.getRow();
        short lastCellNum = row.getLastCellNum();
        String sheetName = context.getWriteSheetHolder().getSheetName();
        if (BooleanUtils.isTrue(context.getHead())) {
            return;
        }
        if (!isTreeReport(sheetName)) {
            return;
        }
        double level = 0;
        if (isSomeReport(sheetName)) {
            level = row.getCell(4).getNumericCellValue();
        } else {
            level = Double.parseDouble(row.getCell(4).getStringCellValue());
        }
        if (!NumberUtils.equals(1, level)) {
            return;
        }
        for (int i = 0; i < lastCellNum; i++) {
            if (styles[i] == null) {
                Workbook workbook = context.getWriteWorkbookHolder().getWorkbook();
                IndexedColorMap indexedColors = ((SXSSFWorkbook) workbook).getXSSFWorkbook().getStylesSource().getIndexedColors();
                XSSFColor color = new XSSFColor(new java.awt.Color(217, 225, 242), indexedColors);
                styles[i] = (XSSFCellStyle) workbook.createCellStyle();
                styles[i].cloneStyleFrom(row.getCell(i).getCellStyle());
                styles[i].setFillForegroundColor(color);
                styles[i].setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            row.getCell(i).setCellStyle(styles[i]);
        }
    }

    private boolean isTreeReport(String sheetName) {
        return true;
    }

    private boolean isSomeReport(String sheetName) {
        return true;
    }
}
