package com.shiyue.jacob;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * @author jeri
 * @version 1.0.0
 * @company www.shiyuesoft.com
 * @date 2019/5/16 14:20
 * @see
 */
public class TestJacob {
    public static void main(String[] args) {
        //word文件路径及名称
        String docPath = "D:\\jacobdemo\\helloworld.docx";
        //html文件路径及名称
        String fileName = "D:\\jacobdemo\\helloworld.html";
        //创建Word对象，启动WINWORD.exe进程
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        //设置用后台隐藏方式打开
        app.setProperty("Visible", new Variant(false));
        //获取操作word的document调用
        Dispatch documents = app.getProperty("Documents").toDispatch();
        //调用打开命令，同时传入word路径
        Dispatch doc = Dispatch.call(documents, "Open", docPath).toDispatch();
        //调用另外为命令，同时传入html的路径
        Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[]{fileName, new Variant(8)}, new int[1]);
        //关闭document对象
        Dispatch.call(doc, "Close", new Variant(0));
        //关闭WINWORD.exe进程
        Dispatch.call(app, "Quit");
        //清空对象
        doc = null;
        app = null;

    /*    String excelPath = "D:\\jacobdemo\\gongzuobu.xlsx";
        String htmlPath = "D:\\jacobdemo\\gongzuobu.html";

        testExcelToHtml(excelPath,htmlPath);*/

        String pptPath = "D:\\jacobdemo\\haha.pptx";
        String htmlPath = "D:\\jacobdemo\\haha.html";
        testPptToHtml(pptPath,htmlPath);
    }

    /**
     *  excel转html
     * @param excelPath
     * @param htmlPath
     */
    public static void testExcelToHtml(String excelPath,String htmlPath){
        ComThread.InitSTA();
        ActiveXComponent app = null;
        System.out.println("-----excel转html开始----");
        app = new ActiveXComponent("Excel.Application");
        app.setProperty("Visible",new Variant(false));
        Dispatch excels = app.getProperty("Workbooks").toDispatch();
        Dispatch excel = Dispatch.invoke(excels, "Open", Dispatch.Method, new Object[]{excelPath, new Variant(false), new Variant(true)}, new int[1]).toDispatch();
        Dispatch.invoke(excel,"SaveAs",Dispatch.Method,new Object[]{htmlPath,new Variant(44)},new int[1]);
        Dispatch.call(excel,"Close",new Variant(false));
        System.out.println("----excel转html结束----");
        ComThread.Release();
    }

    /**
     * ppt转html
     * @param pptPath
     * @param htmlPath
     */
    public static void testPptToHtml(String pptPath,String htmlPath){
            ComThread.InitSTA();
            ActiveXComponent app = new ActiveXComponent("Powerpoint.Application");
        Dispatch dispatch = app.getProperty("Presentations").toDispatch();
        String s2 = pptPath;
        String s3 = htmlPath;
        Dispatch dispatch1 = Dispatch.call(dispatch, "Open", s2, new Variant(-1), new Variant(-1), new Variant(0)).toDispatch();
        Dispatch.call(dispatch1,"SaveAs",s3,new Variant(12));
        Variant variant = new Variant(-1);
        Dispatch.call(dispatch1,"Close");
        app.invoke("Quit",new Variant[0]);
        ComThread.Release();
        ComThread.quitMainSTA();
    }
}
