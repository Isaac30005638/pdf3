package com.data.bkt;

import com.data.util.DocToPdfUtils;
import com.data.util.XmlToDocUtils;

import java.util.HashMap;
import java.util.Map;

/**
 * <p>
 * 字典 Redis 导入
 * </p>
 *
 * @author BMK
 * @Version 1.0
 * @since 2020/11/5 17:03
 */
public class Test {
    public static void main(String[] args) {
        Map<String,String> dataMap = new HashMap<String, String>();

        dataMap.put("a","Hello World!");
        dataMap.put("b","Hello World!");
        dataMap.put("c","Hello World!");
        dataMap.put("d","Hello World!");

        XmlToDocUtils.makeWord("2-测试.xml","2-测试.zip",dataMap);
        DocToPdfUtils.wordToPDF("2-测试.doc","2-测试.pdf");
    }
}
