package cn.huibao.excel;

import com.google.common.collect.Maps;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.Map;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-11
 * Time: 10:09
 * DES:
 */
public class OpenAllExcelFile {

    /**
     * 获取五类文件 的 workbook
     *
     * @param allFilePaths
     * @return
     */
    public static Map<FileNameEnum, XSSFWorkbook> openAllExcel(Map<FileNameEnum, String> allFilePaths) throws Exception {
        Map<FileNameEnum, XSSFWorkbook> result = Maps.newHashMap();

        for (Map.Entry<FileNameEnum, String> entry : allFilePaths.entrySet()) {
            try {
                XSSFWorkbook workbook = new XSSFWorkbook(entry.getValue());
                result.put(entry.getKey(), workbook);
            } catch (IOException e) {
                String errorMsg = "打开文件" + entry.getValue() + " 失败，请检查文件路径是否正确";
                throw new Exception(errorMsg);
            }
        }

        return result;
    }

}
