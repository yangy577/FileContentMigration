package cn.huibao.ui.controller;

import cn.huibao.ui.view.MainFrame;
import cn.huibao.util.CommonUtil;
import cn.huibao.util.PropUtils;
import cn.huibao.util.SwingUtil;

import javax.swing.*;
import java.io.File;
import java.util.Properties;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-09
 * Time: 12:57
 * DES:
 */
public class MainFrameController {

    private MainFrame mainFrame;
    private JLabel msg;
    private JTextField zhouCi;
    private JTextField txtBeginDate;
    private JTextField txtEndDate;
    private JTextField txtFengZhiShiDuanZhuanZhi;
    private JTextField txtFengZhiZhuanZhi;
    private JTextField txtYongDuShiChang;
    private JTextField txtZaoWanGaoFengZhiShu;
    private JTextField txtQuanLuWang;
    private JTextField txtChuGao;
    private JButton btnChangConfig;
    private JButton btnExecute;
    private JTextField txtOut;
    private JPanel mainPanel;

    public MainFrameController() {
        initComponent();
        // 接着处理 button 方法
    }

    // 处理窗体初始化
    public void shwoMainPanelWindow() {
        mainFrame.setVisible(true);
    }

    private void initComponent() {
        mainFrame = new MainFrame();

        mainPanel = mainFrame.getMainPanel();
        msg = mainFrame.getMsg();
        zhouCi = mainFrame.getZhouCi();
        txtBeginDate = mainFrame.getTxtBeginDate();
        txtEndDate = mainFrame.getTxtEndDate();
        txtFengZhiShiDuanZhuanZhi = mainFrame.getTxtFengZhiShiDuanZhuanZhi();
        txtFengZhiZhuanZhi = mainFrame.getTxtFengZhiZhuanZhi();
        txtYongDuShiChang = mainFrame.getTxtYongDuShiChang();
        txtZaoWanGaoFengZhiShu = mainFrame.getTxtZaoWanGaoFengZhiShu();
        txtQuanLuWang = mainFrame.getTxtQuanLuWang();
        txtChuGao = mainFrame.getTxtChuGao();
        btnChangConfig = mainFrame.getBtnChangConfig();
        btnExecute = mainFrame.getBtnExecute();
        txtOut = mainFrame.getTxtOut();

        String initfileResltMsg = initConfigForTxtField();
        if (!CommonUtil.OK_KEY.equals(initfileResltMsg)) {
            msg.setText(SwingUtil.errFontMsg(initfileResltMsg));
        }
    }

    private String initConfigForTxtField() {
        String result = CommonUtil.OK_KEY;
        // 配置文件为方便编辑查看，直接放在与应用 jar 包同级目录下
        String jarPath = System.getProperty("user.dir") + File.separator + CommonUtil.CONFIGURE_NAME;
        Properties configProp;

        try {
            configProp = PropUtils.readPropertiesFile(jarPath);
        } catch (Exception e) {
            result = e.getMessage();
            return result;
        }

        // 加载配置文件中的数值
        if (configProp.containsKey("FROM_FILE_PATH")) {
            txtChuGao.setText(configProp.getProperty("FROM_FILE_PATH").trim());
        }

        if (configProp.containsKey("RESULT_FILE")) {
            txtOut.setText(configProp.getProperty("RESULT_FILE").trim());
        }

        if (configProp.containsKey("FENG_ZHI_SHIDUAN_ZHUANZHI")) {
            txtFengZhiShiDuanZhuanZhi.setText(configProp.getProperty("FENG_ZHI_SHIDUAN_ZHUANZHI").trim());
        }

        if (configProp.containsKey("FENG_ZHI")) {
            txtFengZhiZhuanZhi.setText(configProp.getProperty("FENG_ZHI").trim());
        }

        if (configProp.containsKey("YONG_DU_SHI_CHANG")) {
            txtYongDuShiChang.setText(configProp.getProperty("YONG_DU_SHI_CHANG").trim());
        }

        if (configProp.containsKey("ZAO_WAN_GAO_FENG_ZHISHU")) {
            txtZaoWanGaoFengZhiShu.setText(configProp.getProperty("ZAO_WAN_GAO_FENG_ZHISHU").trim());
        }

        if (configProp.containsKey("JUE_CE_SHU")) {
            txtQuanLuWang.setText(configProp.getProperty("JUE_CE_SHU").trim());
        }

        return result;
    }
}
