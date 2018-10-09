package cn.huibao;

import cn.huibao.ui.controller.MainFrameController;

/**
 * Created with IntelliJ IDEA
 * Created By yy
 * Date: 2018-10-09
 * Time: 12:48
 * DES: 应用启动
 */
public class Runner {

    public static void main(String[] args) {
        MainFrameController mainFrameController = new MainFrameController();
        mainFrameController.shwoMainPanelWindow();
    }
}
