package cn.huibao.util;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-08
 * Time: 15:37
 * DES: 日期工具类
 */
public class DateUtil {

    public final static String YYYYMMDD = "yyyyMMdd";
    public final static String YYYY_MM_DD = "yyyy-MM-dd";

    /**
     * 将字符串日期转为 localdate
     * @param dateStr 字符串日期，"yyyyMMdd"
     * @return LocalDate
     */
    public static LocalDate getLocalDateFromStr(String dateStr) {
        DateTimeFormatter ymd = DateTimeFormatter.ofPattern("yyyyMMdd");
        return LocalDate.parse(dateStr, ymd);
    }

    public static DayOfWeek getDayOfWeekFromStr(String dateStr) {
        LocalDate localDate = getLocalDateFromStr(dateStr);
        return localDate.getDayOfWeek();
    }

    /**
     * 根据传入的日期字符串（yyyyMMdd），期望的星期，判断是否符合
     * @param dateStr 日期字符串
     * @param exportDay 期望
     * @return 是否
     */
    public static boolean verifyDate(String dateStr, DayOfWeek exportDay) {
        LocalDate localDate = DateUtil.getLocalDateFromStr(dateStr);
        return localDate.getDayOfWeek() == exportDay;
    }
}
