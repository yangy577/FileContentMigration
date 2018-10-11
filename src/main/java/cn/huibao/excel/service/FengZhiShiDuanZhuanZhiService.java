package cn.huibao.excel.service;

import cn.huibao.util.CommonUtil;
import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.DayOfWeek;
import java.util.Map;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-11
 * Time: 15:29
 * DES: 处理拥堵时长 excel
 */
public class FengZhiShiDuanZhuanZhiService {

    private static final int SHEET_MONDAY = 2;
    private static final int SHEET_TUESDAY = 3;
    private static final int SHEET_WEDNESDAY = 4;
    private static final int SHEET_THURSDAY = 5;
    private static final int SHEET_FRIDAY = 6;
    private static final int SHEET_SATURDAY = 7;
    private static final int SHEET_SUNDAY = 8;

    /**
     * 获取本周早晚高峰1-5 上周6 日
     */
    public Map<DayOfWeek, String> findFengZhiShiDuanZhuanZhi(XSSFWorkbook yongDuZhiShuWorkbook) {
        XSSFSheet sheet = yongDuZhiShuWorkbook.getSheet(CommonUtil.SHEET_NAME);
        int lastNum = sheet.getLastRowNum();
        Map<DayOfWeek, String> result = Maps.newHashMap();

        // 本周数据
        Row row = sheet.getRow(lastNum);
        result.put(DayOfWeek.MONDAY,
                row.getCell(SHEET_MONDAY) == null ? "" : row.getCell(SHEET_MONDAY).getStringCellValue());
        result.put(DayOfWeek.THURSDAY,
                row.getCell(SHEET_TUESDAY) == null ? "" : row.getCell(SHEET_TUESDAY).getStringCellValue());
        result.put(DayOfWeek.WEDNESDAY,
                row.getCell(SHEET_WEDNESDAY) == null ? "" : row.getCell(SHEET_WEDNESDAY).getStringCellValue());
        result.put(DayOfWeek.THURSDAY,
                row.getCell(SHEET_THURSDAY) == null ? "" : row.getCell(SHEET_THURSDAY).getStringCellValue());
        result.put(DayOfWeek.FRIDAY,
                row.getCell(SHEET_FRIDAY) == null ? "" : row.getCell(SHEET_FRIDAY).getStringCellValue());

        // 上周数据
        Row lastRow = sheet.getRow(lastNum - 1);
        result.put(DayOfWeek.SATURDAY,
                row.getCell(SHEET_SATURDAY) == null ? "" : row.getCell(SHEET_SATURDAY).getStringCellValue());
        result.put(DayOfWeek.SUNDAY,
                row.getCell(SHEET_SUNDAY) == null ? "" : row.getCell(SHEET_SUNDAY).getStringCellValue());

        return result;
    }
}
