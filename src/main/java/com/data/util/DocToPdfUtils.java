package com.data.util;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;

/**
 * <p>
 * 字典 Redis 导入
 * </p>
 *
 * @author BMK
 * @Version 1.0
 * @since 2020/11/5 16:57
 */
public class DocToPdfUtils {
    private static final int wdFormatPDF = 17; // PDF 格式
    public static void wordToPDF(String sfileName, String toFileName) {
        sfileName = "D:\\一些文件\\1-wordtest\\pdf3\\src\\main\\resources\\outword\\"+sfileName;
        toFileName="D:\\一些文件\\1-wordtest\\pdf3\\src\\main\\resources\\outword\\"+toFileName;

        System.out.println("启动 Word...");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", new Variant(false));
            Dispatch docs = app.getProperty("Documents").toDispatch();
            doc = Dispatch.call(docs, "Open", sfileName).toDispatch();
            System.out.println("打开文档..." + sfileName);
            System.out.println("转换文档到 PDF..." + toFileName);
            File tofile = new File(toFileName);
            if (tofile.exists()) {
                tofile.delete();
            }
            Dispatch.call(doc, "SaveAs", toFileName, // FileName
                    wdFormatPDF);
            long end = System.currentTimeMillis();
            System.out.println("转换完成..用时：" + (end - start) + "ms.");

        } catch (Exception e) {
            System.out.println("========Error:文档转换失败：" + e.getMessage());
        } finally {
            Dispatch.call(doc, "Close", false);
            System.out.println("关闭文档");
            if (app != null)
                app.invoke("Quit", new Variant[] {});
        }
        // 如果没有这句话,winword.exe进程将不会关闭
        ComThread.Release();
    }
}
