package cn.huibao.excel.service;

import cn.huibao.util.CommonUtil;
import cn.huibao.util.DateUtil;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.DayOfWeek;
import java.util.List;
import java.util.Map;

public class QuanLuWangService {

    private static final int START_DAY = 1;
    private static final int EDN_DAY = 8;
    private static final String ZAO_QUANUWANG = "zao";
    private static final String WAN_QUANLUWANG = "wan";
    private static final String SIX_QUANLUWANG = "six";
    private static final String EIGHT_QUANLUWANG = "eight";
    private static final String FENG_QUANLUWANG = "feng";
    private static final String FENGZHI_SHIDUAN_QUANLUWANG = "fengzhiShiduan";

    public Map<String, List<Object>> findQuanLuWang(XSSFWorkbook quanLuWangWorkbook) throws Exception {
        Map<String, List<Object>> result = Maps.newHashMap();
        XSSFSheet sheet = quanLuWangWorkbook.getSheet(CommonUtil.SHEET_NAME);

        List<Object> zao = Lists.newArrayList();
        result.put(ZAO_QUANUWANG, zao);
        List<Object> wan = Lists.newArrayList();
        result.put(WAN_QUANLUWANG, wan);
        List<Object> six = Lists.newArrayList();
        result.put(SIX_QUANLUWANG, six);
        List<Object> eight = Lists.newArrayList();
        result.put(EIGHT_QUANLUWANG, eight);
        List<Object> feng = Lists.newArrayList();
        result.put(FENG_QUANLUWANG, feng);
        List<Object> fengzhiShiduan = Lists.newArrayList();
        result.put(FENGZHI_SHIDUAN_QUANLUWANG, fengzhiShiduan);

        for (int i = START_DAY; i < EDN_DAY; i++) {
            XSSFRow row = sheet.getRow(i);
            zao.add(row.getCell(4).getNumericCellValue());
            feng.add(row.getCell(2).getNumericCellValue());
            fengzhiShiduan.add(row.getCell(3).getStringCellValue());
            wan.add(row.getCell(5).getNumericCellValue());
            six.add(row.getCell(6).getNumericCellValue());
            eight.add(row.getCell(7).getNumericCellValue());
        }

        return result;
    }
}
