package in.hopscotch.ExcelReaderV1.utils;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.springframework.util.StringUtils;

public class PoiUtils {

    public static final String XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /**
     * check the row start with startIndex is blank or not
     */
    public static boolean isBlankRow(XSSFRow basicInfoRow, int startIndex) {
        if (null == basicInfoRow)
            return true;
        for (int index = startIndex; index < basicInfoRow.getPhysicalNumberOfCells() - 1; index++) {
            if (!StringUtils.isEmpty(getStringCellValue(basicInfoRow, index)))
                return false;
        }
        return true;
    }

    /**
     * get string cell value
     */
    public static String getStringCellValue(XSSFRow row, int cellnum) {
        XSSFCell cell = row.getCell(cellnum, Row.CREATE_NULL_AS_BLANK);
        cell.setCellType(Cell.CELL_TYPE_STRING);
        return cell.getStringCellValue().trim();
    }

    public static boolean isBlankCell(Cell cell) {
        return StringUtils.isEmpty(getCellValue(cell));
    }

    public static String getCellValue(Cell cell) {
        String ret;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                ret = "";
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                ret = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
                ret = null;
                break;
            case Cell.CELL_TYPE_FORMULA:
                Workbook wb = cell.getSheet().getWorkbook();
                CreationHelper crateHelper = wb.getCreationHelper();
                FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
                ret = getCellValue(evaluator.evaluateInCell(cell));
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date theDate = cell.getDateCellValue();
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
                    ret = simpleDateFormat.format(theDate);
                } else {
                    ret = NumberToTextConverter.toText(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_STRING:
                ret = cell.getRichStringCellValue().getString();
                break;
            default:
                ret = null;
                break;
        }

        return ret != null ? ret.trim() : null;
    }

    public static void setCellValue(int cellNum, XSSFRow row, Object cellValue, CellStyle cellStyle) {
        XSSFCell xssfCell = row.createCell(cellNum);
        if (null != cellValue && !StringUtils.isEmpty(String.valueOf(cellValue))) {
            xssfCell.setCellValue(String.valueOf(cellValue));
        } else {
            xssfCell.setCellValue("");
        }
        xssfCell.setCellStyle(cellStyle);
    }

    public static void setCellValue(int rowNum, int cellNum, XSSFSheet sheet, Object cellValue, CellStyle cellStyle) {
        XSSFCell xssfCell = sheet.getRow(rowNum).createCell(cellNum);
        if (null != cellValue && !StringUtils.isEmpty(String.valueOf(cellValue))) {
            xssfCell.setCellValue(String.valueOf(cellValue));
        } else {
            xssfCell.setCellValue("");
        }
        xssfCell.setCellStyle(cellStyle);
    }

}
