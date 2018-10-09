package cn.huibao.util;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-09
 * Time: 14:44
 * DES:
 */
public class SwingUtil {

    public static String errFontMsg(String msg) {
        String result = "<html><font style='font-size:16px;'>" +
                "错误信息：" +
                "</font><br/>" +
                "<font color='red' style='font-size:12px;'>" +
                msg +
                "</font></html>";
        return result;
    }
}
