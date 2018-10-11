package cn.huibao.excel.service;

import cn.huibao.util.CommonUtil;
import cn.huibao.util.DateUtil;
import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.DayOfWeek;
import java.util.Map;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-11
 * Time: 10:06
 * DES: 处理早晚高峰指数 excel
 */

public class ZaoWanGaoFengZhiShuService {

    private static final int DATE_COL_NUM = 1;

    /**
     * 获取本周早晚高峰1-5 上周6 日
     */
    public Map<DayOfWeek, Row> findZaoWanGaoFengZhiShu(XSSFWorkbook zaowanGaofengZhiShuWorkbook) throws Exception {
        XSSFSheet sheet = zaowanGaofengZhiShuWorkbook.getSheet(CommonUtil.SHEET_NAME);
        int lastNum = sheet.getLastRowNum();
        Map<DayOfWeek, Row> result = Maps.newHashMap();

        // 判断是第一次循环
        boolean isFirst = true;
        int numOfSetResult = 0;

        for (; numOfSetResult < 7; ) {
            XSSFRow currentRow = sheet.getRow(lastNum);
            // 获取日期信息，
            Cell currentCell = currentRow.getCell(DATE_COL_NUM);
            if (currentCell == null) {
                continue;
            }
            String dateStr = currentRow.getCell(DATE_COL_NUM).getStringCellValue();
            DayOfWeek currentDay = DateUtil.getDayOfWeekFromStr(dateStr);

            // 这块的代码仅在开始执行一次
            // 如果是本周的周六日，则忽略
            if (isFirst) {
                if (DayOfWeek.SUNDAY == currentDay) {
                    lastNum = lastNum - 2;
                    isFirst = false;
                    continue;
                } else if (DayOfWeek.SATURDAY == currentDay) {
                    lastNum = lastNum - 1;
                    isFirst = false;
                    continue;
                } else if (DayOfWeek.FRIDAY == currentDay) {
                    isFirst = false;
                } else {
                    throw new Exception("早晚高峰指数.xlsx 中数据不完整(最后一行数据日期是" + currentDay + ")");
                }
            }

            if (result.containsKey(currentDay)) {
                throw new Exception("早晚高峰指数.xlsx 中,第" + lastNum + "行附近数据有误，日期不连续");
            }
            result.put(currentDay, currentRow);
            numOfSetResult++;
            lastNum--;
        }

        return result;
    }

}
