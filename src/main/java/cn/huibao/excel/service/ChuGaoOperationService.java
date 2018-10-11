package cn.huibao.excel.service;

import cn.huibao.util.CommonUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.DayOfWeek;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-11
 * Time: 14:25
 * DES: 指数预测 sheet 操作类
 */
public class ChuGaoOperationService {

    private static final int START_ROW_NUM = 13;
    private static final int SHIFT_ROW_NUM = 6;

    private static final int SHEET_MONDAY = 2;
    private static final int SHEET_TUESDAY = 3;
    private static final int SHEET_WEDNESDAY = 4;
    private static final int SHEET_THURSDAY = 5;
    private static final int SHEET_FRIDAY = 6;
    private static final int SHEET_SATURDAY = 7;
    private static final int SHEET_SUNDAY = 8;

    // 早高峰所在行
    private static final int ROW_NUM_OF_ZAO = 13;
    // 上周早高峰所在行
    private static final int ROW_NUM_OF_ZAO_LAST = 19;
    // 晚高峰所在行
    private static final int ROW_NUM_OF_WAN = 14;
    // 上周晚高峰所在行
    private static final int ROW_NUM_OF_WAN_LAST = 20;
    // 峰值所在行
    private static final int ROW_NUM_OF_FENGZHI = 15;
    // 上周峰值所在行
    private static final int ROW_NUM_OF_FENGZHI_LAST = 21;
    // 峰值时段转置所在行
    public static final int ROW_NUM_OF_FENGZHISHIDUAN = 16;
    // 上周峰值时段转置所在行
    public static final int ROW_NUM_OF_FENGZHISHIDUAN_LAST = 22;
    // 6以上累计时长所在行
    private static final int ROW_NUM_OF_SIX = 17;
    // 上周6以上累计时长
    private static final int ROW_NUM_OF_SIX_LAST = 23;
    // 8以上累计时长所在行
    private static final int ROW_NUM_OF_EIGHT = 18;
    // 上周8以上累计时长
    private static final int ROW_NUM_OF_EIGHT_LAST = 24;

    // 在 sheet1-早晚高峰指数excel 中的列数
    private static final int ZAOWANFENG_SHEET_ZAO_NUM = 5;
    private static final int ZAOWANFENG_SHEET_WAN_NUM = 4;
    private static final int ZAOWANFENG_SHEET_FENG_NUM = 2;

    // 在 sheet1-拥堵指数excel 中的列数 6
    private static final int YONGDUSHICHANG_SHEET_SIX = 2;
    // 在 sheet1-拥堵指数excel 中的列数 8
    private static final int YONGDUSHICHANG_SHEET_EIGHT = 3;

    // 决策树中用到的
    private static final String ZAO_QUANUWANG = "zao";
    private static final String WAN_QUANLUWANG = "wan";
    private static final String SIX_QUANLUWANG = "six";
    private static final String EIGHT_QUANLUWANG = "eight";
    private static final String FENG_QUANLUWANG = "feng";
    private static final String FENGZHI_SHIDUAN_QUANLUWANG = "fengzhiShiduan";


    private ZaoWanGaoFengZhiShuService zaoWanGaoFengZhiShuService;
    private FengZhiShiDuanZhuanZhiService fengZhiShiDuanZhuanZhiService;
    private YongDuShiChangService yongDuShiChangService;
    private QuanLuWangService quanLuWangService;

    public ChuGaoOperationService() {
        this.zaoWanGaoFengZhiShuService = new ZaoWanGaoFengZhiShuService();
        this.fengZhiShiDuanZhuanZhiService = new FengZhiShiDuanZhuanZhiService();
        this.yongDuShiChangService = new YongDuShiChangService();
        this.quanLuWangService = new QuanLuWangService();
    }

    // 获取早晚高峰指数
    //Map<DayOfWeek, Row> zaoWanGaoFengZhiShu = zaoWanGaoFengZhiShuService.findZaoWanGaoFengZhiShu(zaowangaofengWorkbook);
    // 获取 峰值时段转置表.xlsx 中的峰值
    //Map<DayOfWeek, String> fengZhiShiDuanZhuanZhi = fengZhiShiDuanZhuanZhiService.findFengZhiShiDuanZhuanZhi()
    //
    public void operations(XSSFWorkbook workbook,
                           XSSFWorkbook zaowangaofengWorkbook,
                           XSSFWorkbook yongDuZhiShuWorkBook,
                           XSSFWorkbook fengZhiShiDuanZhuanZhiWorkbook,
                           XSSFWorkbook quanLuWangWorkbook,
                           String zhou, String benginDate, String endDate) throws Exception {
        XSSFSheet sheet = workbook.getSheet(CommonUtil.SHEET_ZHI_SHU_YUCE);
        if (sheet == null) {
            throw new Exception("无法获取" + CommonUtil.SHEET_ZHI_SHU_YUCE + ", 请检查初稿中是否有该Sheet");
        }

        // 获取数据
        Map<DayOfWeek, Row> yongDuShiChang = yongDuShiChangService.findYongDuShiChang(yongDuZhiShuWorkBook);
        Map<DayOfWeek, String> fengZhiShiDuan = fengZhiShiDuanZhuanZhiService.findFengZhiShiDuanZhuanZhi(fengZhiShiDuanZhuanZhiWorkbook);
        Map<DayOfWeek, Row> zaoWanGaoFengZhiShu = zaoWanGaoFengZhiShuService.findZaoWanGaoFengZhiShu(zaowangaofengWorkbook);
        Map<String, List<Object>> quanluwang = quanLuWangService.findQuanLuWang(quanLuWangWorkbook);

        /* 开始处理 指数预测Sheet */
        // 插入新一周
        insertNewWeekZhiShu(workbook, sheet, zhou, benginDate, endDate);

        // 插入早高峰
        insertZaoGaoFeng(sheet, zaoWanGaoFengZhiShu);
        // 插入晚高峰
        insertWanGaoFeng(sheet, zaoWanGaoFengZhiShu);
        // 插入峰值
        insertFeng(sheet, zaoWanGaoFengZhiShu);

        // 插入峰值时段
        insertFengZhiShiDuan(sheet, fengZhiShiDuan);

        // 插入6以上累计时长
        insertYongDuSix(sheet, yongDuShiChang);

        // 插入8以上累计时长
        insertYongDuEight(sheet, yongDuShiChang);

        // 插入决策树
        insertQuanLuWang(sheet, quanluwang);
        /* 结束处理 指数预测Sheet */

        /* 开始处理 早晚高峰指数sheet */
        XSSFSheet zwgfzsSheet = workbook.getSheet(CommonUtil.SHEET_ZAO_WAN_GAO_FENG);
        if (zwgfzsSheet == null) {
            throw new Exception("无法获取" + CommonUtil.SHEET_ZHI_SHU_YUCE + ", 请检查初稿中是否有该Sheet");
        }
        insertSheetZaoWanGaoFeng(zwgfzsSheet, zaoWanGaoFengZhiShu);
        /* 结束处理 早晚高峰指数sheet */

        /* 开始处理 拥堵时长sheet */
        XSSFSheet ydscSheet = workbook.getSheet(CommonUtil.SHEET_YONG_DU);
        if (ydscSheet == null) {
            throw new Exception("无法获取" + CommonUtil.SHEET_YONG_DU + ", 请检查初稿中是否有该Sheet");
        }
        insertSheetYongDuShiChang(ydscSheet, yongDuShiChang);
        /* 结束处理 拥堵时长sheet */
    }


    /**
     * 在14行的位置插入复制6行并插入
     * 插入新一周的
     */
    private void insertNewWeekZhiShu(XSSFWorkbook workbook, XSSFSheet zhiShuYuCeSheet, String zhou, String benginDate, String endDate) {
        zhiShuYuCeSheet.shiftRows(START_ROW_NUM, zhiShuYuCeSheet.getLastRowNum(), SHIFT_ROW_NUM);

        for (int i = (START_ROW_NUM + SHIFT_ROW_NUM), j = START_ROW_NUM; i < (START_ROW_NUM + SHIFT_ROW_NUM + SHIFT_ROW_NUM); i++, j++) {
            copyRow(workbook, i, j, true, zhiShuYuCeSheet, zhou, benginDate, endDate);
        }
    }

    private void insertFeng(XSSFSheet sheet, Map<DayOfWeek, Row> zaoWanGaoFengZhiShu) {
        // 峰值 15 行
        commonInsertFeng(sheet, zaoWanGaoFengZhiShu, ROW_NUM_OF_FENGZHI, ZAOWANFENG_SHEET_FENG_NUM, ROW_NUM_OF_FENGZHI_LAST);
    }

    private void insertWanGaoFeng(XSSFSheet sheet, Map<DayOfWeek, Row> zaoWanGaoFengZhiShu) {
        // 晚高峰 14 行
        commonInsertFeng(sheet, zaoWanGaoFengZhiShu, ROW_NUM_OF_WAN, ZAOWANFENG_SHEET_WAN_NUM, ROW_NUM_OF_WAN_LAST);
    }

    // 插入早高峰
    private void insertZaoGaoFeng(XSSFSheet sheet, Map<DayOfWeek, Row> zaoWanGaoFengZhiShu) {
        // 早高峰 13 行
        commonInsertFeng(sheet, zaoWanGaoFengZhiShu, ROW_NUM_OF_ZAO, ZAOWANFENG_SHEET_ZAO_NUM, ROW_NUM_OF_ZAO_LAST);
    }

    // 插入峰值时段
    private void insertFengZhiShiDuan(XSSFSheet sheet, Map<DayOfWeek, String> fengZhiShiDuan) {
        // 峰值时段 16 行
        //commonInsertFeng(sheet, fengZhiShiDuan, );
        XSSFRow row = sheet.getRow(ROW_NUM_OF_FENGZHISHIDUAN);
        row.getCell(SHEET_MONDAY).setCellValue(fengZhiShiDuan.get(DayOfWeek.MONDAY));
        row.getCell(SHEET_TUESDAY).setCellValue(fengZhiShiDuan.get(DayOfWeek.TUESDAY));
        row.getCell(SHEET_WEDNESDAY).setCellValue(fengZhiShiDuan.get(DayOfWeek.WEDNESDAY));
        row.getCell(SHEET_THURSDAY).setCellValue(fengZhiShiDuan.get(DayOfWeek.THURSDAY));
        row.getCell(SHEET_FRIDAY).setCellValue(fengZhiShiDuan.get(DayOfWeek.FRIDAY));
        XSSFRow lastRow = sheet.getRow(ROW_NUM_OF_FENGZHISHIDUAN_LAST);
        row.getCell(SHEET_SATURDAY).setCellValue(fengZhiShiDuan.get(DayOfWeek.SATURDAY));
        row.getCell(SHEET_SUNDAY).setCellValue(fengZhiShiDuan.get(DayOfWeek.SUNDAY));
    }

    // 插入6以上累计时长
    private void insertYongDuSix(XSSFSheet sheet, Map<DayOfWeek, Row> yongDuShiChang) {
        commonInsertFeng(sheet, yongDuShiChang, ROW_NUM_OF_SIX, YONGDUSHICHANG_SHEET_SIX, ROW_NUM_OF_SIX_LAST);
    }

    // 插入8以上累计时长
    private void insertYongDuEight(XSSFSheet sheet, Map<DayOfWeek, Row> yongDuShiChang) {
        commonInsertFeng(sheet, yongDuShiChang, ROW_NUM_OF_EIGHT, YONGDUSHICHANG_SHEET_EIGHT, ROW_NUM_OF_EIGHT_LAST);
    }

    // 将 早晚高峰指数excel 中周一-周五，上周六-周日数据插入 初稿的早晚高峰Sheet中
    private void insertSheetZaoWanGaoFeng(XSSFSheet zwgfzsSheet, Map<DayOfWeek, Row> zaoWanGaoFengZhiShu) {
        int lastNum = zwgfzsSheet.getLastRowNum();

        // 获取几种样式
        XSSFRow styleRow = zwgfzsSheet.getRow(zwgfzsSheet.getLastRowNum() - 3);
        CellStyle firstCellStyle = styleRow.getCell(0).getCellStyle();
        CellStyle normalCellStyle = styleRow.getCell(1).getCellStyle();
        CellStyle workingDayCellStyle = styleRow.getCell(4).getCellStyle();

        for (int i = 0; i < 7; i++) {
            XSSFRow row = zwgfzsSheet.createRow(zwgfzsSheet.getLastRowNum() + 1);
            row.setHeight(styleRow.getHeight());
            XSSFRow fromRow = getZaoWanGaoFengZhiShuFromMap(zaoWanGaoFengZhiShu, i);
            for (int j = 0; j < 8; j++) {
                Cell cell = row.createCell(j);
                if (j == 0) {
                    cell.setCellStyle(firstCellStyle);
                    int num = (int) (zwgfzsSheet.getRow(lastNum).getCell(0).getNumericCellValue() + 1);
                    cell.setCellValue(num);
                    continue;
                }
                if (j == 5 || j == 6) {
                    cell.setCellStyle(workingDayCellStyle);
                    cell.setCellValue(fromRow.getCell(j).getNumericCellValue());
                    continue;
                }
                if (normalCellStyle != null) {
                    cell.setCellStyle(normalCellStyle);
                    if (j == 1) {
                        cell.setCellValue(fromRow.getCell(j).getStringCellValue());
                    } else {
                        cell.setCellValue(fromRow.getCell(j).getNumericCellValue());
                    }
                }
            }
        }
    }

    // 将 拥堵时长excel 中周一-周五，上周六-周日数据插入 初稿的拥堵时长Sheet中
    private void insertSheetYongDuShiChang(XSSFSheet ydscSheet, Map<DayOfWeek, Row> yongDuShiChang) {
        int lastNum = ydscSheet.getLastRowNum();

        // 获取几种样式
        XSSFRow styleRow = ydscSheet.getRow(ydscSheet.getLastRowNum() - 3);
        CellStyle firstCellStyle = styleRow.getCell(0).getCellStyle();
        CellStyle normalCellStyle = styleRow.getCell(1).getCellStyle();

        for (int i = 0; i < 7; i++) {
            XSSFRow row = ydscSheet.createRow(ydscSheet.getLastRowNum() + 1);
            row.setHeight(styleRow.getHeight());
            XSSFRow fromRow = getZaoWanGaoFengZhiShuFromMap(yongDuShiChang, i);
            for (int j = 0; j < 4; j++) {
                Cell cell = row.createCell(j);
                if (j == 0) {
                    cell.setCellStyle(firstCellStyle);
                    int num = (int) (ydscSheet.getRow(lastNum).getCell(0).getNumericCellValue() + 1);
                    cell.setCellValue(num);
                    continue;
                }
                if (normalCellStyle != null) {
                    cell.setCellStyle(normalCellStyle);
                    if (j == 1) {
                        cell.setCellValue(fromRow.getCell(j).getStringCellValue());
                    } else {
                        cell.setCellValue(fromRow.getCell(j).getNumericCellValue());
                    }
                }
            }
        }
    }

    // 插入决策树 -
    private void insertQuanLuWang(XSSFSheet sheet, Map<String, List<Object>> juceshu) {
        for (int colNum = 2; colNum < 9; colNum++) {
            sheet.getRow(1).getCell(colNum).setCellValue((Double) juceshu.get(ZAO_QUANUWANG).get(colNum - 2));
            sheet.getRow(2).getCell(colNum).setCellValue((Double) juceshu.get(WAN_QUANLUWANG).get(colNum - 2));
            sheet.getRow(3).getCell(colNum).setCellValue((Double) juceshu.get(FENG_QUANLUWANG).get(colNum - 2));
            sheet.getRow(4).getCell(colNum).setCellValue((String) juceshu.get(FENGZHI_SHIDUAN_QUANLUWANG).get(colNum - 2));
            sheet.getRow(5).getCell(colNum).setCellValue((Double) juceshu.get(SIX_QUANLUWANG).get(colNum - 2));
            sheet.getRow(6).getCell(colNum).setCellValue((Double) juceshu.get(EIGHT_QUANLUWANG).get(colNum - 2));
        }
    }

    private void commonInsertFeng(XSSFSheet sheet, Map<DayOfWeek, Row> zaoWanGaoFengZhiShu, int rowNum, int zaowanfengNumInSheet, int rowNumLast) {
        XSSFRow row = sheet.getRow(rowNum);

        row.getCell(SHEET_MONDAY).setCellValue(
                zaoWanGaoFengZhiShu.get(DayOfWeek.MONDAY).getCell(zaowanfengNumInSheet).getNumericCellValue()
        );
        row.getCell(SHEET_TUESDAY).setCellValue(
                zaoWanGaoFengZhiShu.get(DayOfWeek.TUESDAY).getCell(zaowanfengNumInSheet).getNumericCellValue()
        );
        row.getCell(SHEET_WEDNESDAY).setCellValue(
                zaoWanGaoFengZhiShu.get(DayOfWeek.WEDNESDAY).getCell(zaowanfengNumInSheet).getNumericCellValue()
        );
        row.getCell(SHEET_THURSDAY).setCellValue(
                zaoWanGaoFengZhiShu.get(DayOfWeek.THURSDAY).getCell(zaowanfengNumInSheet).getNumericCellValue()
        );
        row.getCell(SHEET_FRIDAY).setCellValue(
                zaoWanGaoFengZhiShu.get(DayOfWeek.FRIDAY).getCell(zaowanfengNumInSheet).getNumericCellValue()
        );

        XSSFRow lastRow = sheet.getRow(rowNumLast);
        lastRow.getCell(SHEET_SATURDAY).setCellValue(
                zaoWanGaoFengZhiShu.get(DayOfWeek.MONDAY).getCell(zaowanfengNumInSheet).getNumericCellValue()
        );
        lastRow.getCell(SHEET_SUNDAY).setCellValue(
                zaoWanGaoFengZhiShu.get(DayOfWeek.SUNDAY).getCell(zaowanfengNumInSheet).getNumericCellValue()
        );
    }

    /**
     * 行复制功能
     */
    public static void copyRow(Workbook wb, int fromRowNum, int toRowNum, boolean copyValueFlag, XSSFSheet sheet,
                               String zhou, String beginDate, String endDate) {
        Row toRow = sheet.getRow(toRowNum);
        if (toRow == null) {
            toRow = sheet.createRow(toRowNum);
        }
        Row fromRow = sheet.getRow(fromRowNum);
        toRow.setHeight(fromRow.getHeight());
        for (Iterator cellIt = fromRow.cellIterator(); cellIt.hasNext(); ) {
            Cell tmpCell = (Cell) cellIt.next();
            Cell newCell = toRow.createCell(tmpCell.getColumnIndex());
            copyCell(wb, tmpCell, newCell, copyValueFlag, zhou, beginDate, endDate);
        }
        Sheet worksheet = fromRow.getSheet();
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == fromRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(toRow.getRowNum(), (toRow.getRowNum() +
                        (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())), cellRangeAddress
                        .getFirstColumn(), cellRangeAddress.getLastColumn());
                worksheet.addMergedRegionUnsafe(newCellRangeAddress);
            }
        }
    }

    private XSSFRow getZaoWanGaoFengZhiShuFromMap(Map<DayOfWeek, Row> zaoWanGaoFengZhiShu, int i) {
        if (i == 0) {
            return (XSSFRow) zaoWanGaoFengZhiShu.get(DayOfWeek.SATURDAY);
        }
        if (i == 1) {
            return (XSSFRow) zaoWanGaoFengZhiShu.get(DayOfWeek.SUNDAY);
        }
        if (i == 2) {
            return (XSSFRow) zaoWanGaoFengZhiShu.get(DayOfWeek.MONDAY);
        }
        if (i == 3) {
            return (XSSFRow) zaoWanGaoFengZhiShu.get(DayOfWeek.TUESDAY);
        }
        if (i == 4) {
            return (XSSFRow) zaoWanGaoFengZhiShu.get(DayOfWeek.WEDNESDAY);
        }
        if (i == 5) {
            return (XSSFRow) zaoWanGaoFengZhiShu.get(DayOfWeek.THURSDAY);
        }
        if (i == 6) {
            return (XSSFRow) zaoWanGaoFengZhiShu.get(DayOfWeek.FRIDAY);
        }
        return null;
    }

    /**
     * 复制单元格
     *
     * @param srcCell
     * @param distCell
     * @param copyValueFlag true则连同cell的内容一起复制
     */
    public static void copyCell(Workbook wb, Cell srcCell, Cell distCell, boolean copyValueFlag,
                                String zhou, String beginDate, String endDate) {
        CellStyle newStyle = wb.createCellStyle();
        CellStyle srcStyle = srcCell.getCellStyle();
        newStyle.cloneStyleFrom(srcStyle);
        newStyle.setFont(wb.getFontAt(srcStyle.getFontIndex()));
        //样式
        distCell.setCellStyle(newStyle);
        //评论
        if (srcCell.getCellComment() != null) {
            distCell.setCellComment(srcCell.getCellComment());
        }
        // 不同数据类型处理
        CellType srcCellType = srcCell.getCellTypeEnum();
        distCell.setCellType(srcCellType);
        if (copyValueFlag) {
            if (srcCellType == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(srcCell)) {
                    distCell.setCellValue(srcCell.getDateCellValue());
                } else {
                    distCell.setCellValue(srcCell.getNumericCellValue());
                }
            } else if (srcCellType == CellType.STRING) {
                // 复制的第一个column是 ： 周次+本周起止时间
                if (srcCell.getColumnIndex() == 0) {
                    StringBuilder sb = new StringBuilder();
                    sb.append(
                            zhou).append("周（").append(beginDate).append("-").append(endDate).append("）");
                    distCell.setCellValue(sb.toString());
                } else {
                    distCell.setCellValue(srcCell.getRichStringCellValue());
                }
            } else if (srcCellType == CellType.BLANK) {

            } else if (srcCellType == CellType.BOOLEAN) {
                distCell.setCellValue(srcCell.getBooleanCellValue());
            } else if (srcCellType == CellType.ERROR) {
                distCell.setCellErrorValue(srcCell.getErrorCellValue());
            } else if (srcCellType == CellType.FORMULA) {
                distCell.setCellFormula(testReplace(srcCell.getCellFormula(), SHIFT_ROW_NUM));
            } else {
            }
        }
    }

    public static String testReplace(String sourceStr, int shiftNum) {
        StringBuilder sb = new StringBuilder();
        String digitsStr = "";

        for (int i = 0; i < sourceStr.length(); i++) {
            if (Character.isDigit(sourceStr.charAt(i))) {
                digitsStr += sourceStr.charAt(i);
                if (i != sourceStr.length() - 1) {
                    continue;
                }
            }

            if (digitsStr.length() > 0) {
                int num = Integer.valueOf(digitsStr) - shiftNum;
                sb.append(num);
                digitsStr = "";
            }
            if (!Character.isDigit(sourceStr.charAt(i))) {
                sb.append(sourceStr.charAt(i));
            }

        }

        return sb.toString();
    }
}
