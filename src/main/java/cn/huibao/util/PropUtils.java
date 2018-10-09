package cn.huibao.util;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.util.Properties;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-08
 * Time: 16:35
 * DES:
 */
public class PropUtils {
    /**
     * 读取prop 文件
     *
     * @param path 文件路径
     * @return prop
     */
    public static Properties readPropertiesFile(String path) throws IOException {
        Properties prop = new Properties();
        InputStream in = null;
        try {
            in = new BufferedInputStream(new FileInputStream(path));
            prop.load(new InputStreamReader(in, "utf-8"));
            for (String key : prop.stringPropertyNames()) {
                System.out.println(key + ":" + prop.getProperty(key));
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());
            throw e;
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    System.out.println(e.getMessage());
                    throw e;
                }
            }
        }
        return prop;
    }

    /**
     * 写Properties文件
     */
    public static void writePropertiesFile(Properties prop, String path) {
        FileOutputStream oFile = null;
        try {
            //保存属性到b.properties文件
            oFile = new FileOutputStream(path, false);//true表示追加打开,false每次都是清空再重写
            //prop.store(oFile, "此参数是保存生成properties文件中第一行的注释说明文字");//这个会两个地方乱码
            //prop.store(new OutputStreamWriter(oFile, "utf-8"), "汉字乱码");//这个就是生成的properties文件中第一行的注释文字乱码
            prop.store(new OutputStreamWriter(oFile, "utf-8"), "配置文件");
        } catch (Exception e) {
            System.out.println(e.getMessage());
        } finally {
            if (oFile != null) {
                try {
                    oFile.flush();
                    oFile.close();
                    oFile.flush();
                } catch (IOException e) {
                    System.out.println(e.getMessage());
                }
            }
        }
    }
}
