package com.data.util;


import freemarker.template.Configuration;
import freemarker.template.Template;


import java.io.*;
import java.util.Map;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 *
 * @author BMK
 * @Version 1.0
 * @since 2020/11/5 15:47
 */
public class XmlToDocUtils {

    /**
     *利用freemarker动态插入数据
     * @param xmlName   从zip文件中解压出来的xml文件名称，更改为和zip同名称
     * @param zipName   压缩成的zip文件
     * @param dataMap   需要传入的数据
     * @throws Exception
     */
    public static void makeWord(String xmlName,String zipName,Map dataMap) {
        //输入所需要的zip，xml文件所在目录
        String inDirectory ="D:\\一些文件\\1-wordtest\\pdf3\\src\\main\\resources\\inword\\";
        //输出目录
        String outDirectory = "D:\\一些文件\\1-wordtest\\pdf3\\src\\main\\resources\\outword\\";
        //初始化配置文件
        Configuration configuration = new Configuration();
        try {
        //加载文件
        configuration.setDirectoryForTemplateLoading(new File(inDirectory));
        // 加载zip解压的xml模板
        Template template = configuration.getTemplate(xmlName);
        
        //命名更改元素后的xml名称
        String xmlOutName = xmlName.split("\\.")[0]+"update.xml";

        // 指定更改元素后输出的xml名称,输出到outword文件夹
        String outFilePath = outDirectory+"/"+xmlOutName;
        File docFile = new File(outFilePath);
        FileOutputStream fos = new FileOutputStream(docFile);
        Writer out = new BufferedWriter(new OutputStreamWriter(fos),10240);

        //利用freemarker进行替换输出

        template.process(dataMap,out);
        if(out != null){
            out.close();
        }


        //命名输出的doc文件
        String fileName = xmlName.split("\\.")[0]+".doc";

            //绑定输入的zip文件名称
            ZipInputStream zipInputStream = ZipUtils.wrapZipInputStream(new FileInputStream(new File(inDirectory+zipName)));
            //绑定输出的doc文件名称
            ZipOutputStream zipOutputStream = ZipUtils.wrapZipOutputStream(new FileOutputStream(new File(outDirectory+fileName)));
            //对zip中对应的docement文件进行替换并解压输出
            String itemname = "word/document.xml";
            ZipUtils.replaceItem(zipInputStream, zipOutputStream, itemname, new FileInputStream(new File(outDirectory+xmlOutName)));
            System.out.println("success");

        } catch (Exception e) {
            System.out.println(e.toString());
        }




    }


}
