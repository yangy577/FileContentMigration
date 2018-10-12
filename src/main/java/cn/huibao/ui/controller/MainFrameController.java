package cn.huibao.ui.controller;

import cn.huibao.excel.service.ChuGaoOperationService;
import cn.huibao.ui.view.MainFrame;
import cn.huibao.util.CommonUtil;
import cn.huibao.util.PropUtils;
import cn.huibao.util.SwingUtil;
import com.google.common.base.Strings;
import com.google.common.collect.Maps;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Map;
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

    private ChuGaoOperationService operationService;

    private Map<String, String> filePathMap = Maps.newHashMap();

    public MainFrameController() {
        operationService = new ChuGaoOperationService();

        initComponent();
        // 接着处理 button 方法
        initListeners();
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

    private void initListeners() {
        btnExecute.addActionListener((new BtnExecuteLister()));
        btnChangConfig.addActionListener(new BtnChangConfigLister());
    }

    private class BtnExecuteLister implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            try {
                btnExecute.setEnabled(false);
                // 周次，两个时间必须填写
                String returnMsg = verifyZhouci();
                if (!CommonUtil.OK_KEY.equals(returnMsg)) {
                    msg.setText(SwingUtil.errFontMsg(returnMsg));
                    btnExecute.setEnabled(true);
                    return;
                }
                returnMsg = writeToConfig();
                if (!CommonUtil.OK_KEY.equals(returnMsg)) {
                    msg.setText(SwingUtil.errFontMsg(returnMsg));
                    btnExecute.setEnabled(true);
                    return;
                }
            } catch (Exception e1) {
                msg.setText(e1.getMessage());
                btnExecute.setEnabled(true);
            }

            try {
                operationService.operations(filePathMap, zhouCi.getText().trim(), txtBeginDate.getText().trim(), txtEndDate.getText().trim());
                btnExecute.setEnabled(true);
                msg.setText("已完成");
            } catch (Exception e1) {
                msg.setText(SwingUtil.errFontMsg(e1.getMessage()));
                btnExecute.setEnabled(true);
            }
        }
    }

    private class BtnChangConfigLister implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            try {
                btnExecute.setEnabled(false);
                String returnMsg = writeToConfig();
                if (!CommonUtil.OK_KEY.equals(returnMsg)) {
                    msg.setText(SwingUtil.errFontMsg(returnMsg));
                    btnExecute.setEnabled(true);
                    return;
                }
            } catch (Exception e1) {
                msg.setText(e1.getMessage());
                btnExecute.setEnabled(true);
            }
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

    // 校验并返回错误信息
    private String writeToConfig() {
        String jarPath = System.getProperty("user.dir") + File.separator + CommonUtil.CONFIGURE_NAME;
        Properties configProp;
        try {
            configProp = PropUtils.readPropertiesFile(jarPath);
        } catch (Exception e) {
            return e.toString();
        }

        String errMsg = "";
        if (Strings.isNullOrEmpty(txtChuGao.getText().trim())) {
            errMsg += "原始初稿路径信息缺失，";
        }

        if (Strings.isNullOrEmpty(txtOut.getText().trim())) {
            errMsg += "生成文件路径信息缺失，";
        }

        if (Strings.isNullOrEmpty(txtFengZhiShiDuanZhuanZhi.getText().trim())) {
            errMsg += "峰值时段转置表路径信息缺失，";
        }

        if (Strings.isNullOrEmpty(txtFengZhiZhuanZhi.getText().trim())) {
            errMsg += "峰值转置表路径信息缺失，";
        }

        if (Strings.isNullOrEmpty(txtYongDuShiChang.getText().trim())) {
            errMsg += "拥堵时长路径信息缺失，";
        }

        if (Strings.isNullOrEmpty(txtZaoWanGaoFengZhiShu.getText().trim())) {
            errMsg += "早晚高峰指数路径信息缺失，";
        }

        if (Strings.isNullOrEmpty(txtQuanLuWang.getText().trim())) {
            errMsg += "全路网高峰指数_峰值_峰值时段表路径信息缺失，";
        }

        if (errMsg.length() > 0) {
            errMsg += "请先配置路径信息";
            return errMsg;
        } else {
            filePathMap.put(CommonUtil.NAME_FENG_ZHI_SHI_DUAN_ZHUAN_ZHI, txtFengZhiShiDuanZhuanZhi.getText().trim());
            filePathMap.put(CommonUtil.NAME_FENG_ZHI_ZHUAN_ZHI, txtFengZhiZhuanZhi.getText().trim());
            filePathMap.put(CommonUtil.NAME_QUAN_LU_WANG, txtQuanLuWang.getText().trim());
            filePathMap.put(CommonUtil.NAME_YONG_DU_SHI_CHANG, txtYongDuShiChang.getText().trim());
            filePathMap.put(CommonUtil.NAME_ZAO_WAN_GAO_FENG_ZHI_SHU, txtZaoWanGaoFengZhiShu.getText().trim());
            filePathMap.put(CommonUtil.NAME_FROM, txtChuGao.getText().trim());
            filePathMap.put(CommonUtil.NAME_TO, txtOut.getText().trim());

            configProp.setProperty("ZAO_WAN_GAO_FENG_ZHISHU", txtZaoWanGaoFengZhiShu.getText().trim());
            configProp.setProperty("JUE_CE_SHU", txtQuanLuWang.getText().trim());
            configProp.setProperty("RESULT_FILE", txtOut.getText().trim());
            configProp.setProperty("FROM_FILE_PATH", txtChuGao.getText().trim());
            configProp.setProperty("YONG_DU_SHI_CHANG", txtYongDuShiChang.getText().trim());
            configProp.setProperty("FENG_ZHI_SHIDUAN_ZHUANZHI", txtFengZhiShiDuanZhuanZhi.getText().trim());
            configProp.setProperty("FENG_ZHI", txtFengZhiZhuanZhi.getText().trim());

            PropUtils.writePropertiesFile(configProp, jarPath);
            return CommonUtil.OK_KEY;
        }
    }


    // 周次、开始日期、结束日期 必须填写
    private String verifyZhouci() {
        String errMsg = "";

        if (Strings.isNullOrEmpty(zhouCi.getText().trim())) {
            errMsg += "周次、";
        }
        if (Strings.isNullOrEmpty(txtBeginDate.getText().trim())) {
            errMsg += "开始日期、";
        }
        if (Strings.isNullOrEmpty(txtEndDate.getText().trim())) {
            errMsg += "结束日期";
        }
        if (errMsg.length() > 0) {
            errMsg += "必须填写；";
            return errMsg;
        }
        return CommonUtil.OK_KEY;
    }

}
