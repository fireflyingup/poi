package com.poi.study.service;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.DocxRenderData;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

@Service
public class PoiService{

    private static long start = 0L;

    public void test1() throws Exception{
        start = System.currentTimeMillis();
        List<Father> list = initData();
        printTime("生产数据：");
        File templateFile = new File("C:\\Users\\13015\\Desktop\\temp\\template.docx");
        XWPFTemplate compile = XWPFTemplate.compile(templateFile);
        printTime("编译模板：");
        Map<String, Object> map = new HashMap<>();
        map.put("father", list);
        XWPFTemplate render = compile.render(map);
        printTime("渲染数据：");
//        render.writeToFile("C:\\Users\\13015\\Desktop\\temp\\result.docx");
        printTime("写入数据：");
    }
//    生产数据：1ms
//    合并数据：159314ms
//    写入数据：1145ms
//    生产数据：1ms
//    合并数据：152995ms
//    写入数据：1176ms
    public void test2() throws Exception {
        ZipSecureFile.setMinInflateRatio(0d);
        System.out.println("-----------------------------------");
        start = System.currentTimeMillis();
        List<Father> list = initData();
        printTime("生产数据：");
        NiceXWPFDocument niceXWPFDocument = null;
        List<NiceXWPFDocument> documentList = new LinkedList<>();
        XWPFTemplate xwpfTemplate;
        XWPFTemplate xwpfTemplate1 = XWPFTemplate.compile("C:\\Users\\13015\\Desktop\\temp\\template2.docx");
        XWPFTemplate render = null;
        DocxRenderData docxRenderData;
        Map<String, Object> map = new HashMap<>();
        int i = 0;
        for (Father father : list) {
            xwpfTemplate = XWPFTemplate.compile("C:\\Users\\13015\\Desktop\\temp\\template1.docx");
            docxRenderData = new DocxRenderData(new File("C:\\Users\\13015\\Desktop\\temp\\template2.docx"), father.getChildren());
            map.put("list", docxRenderData);
            map.put("country", father.getCountry());
            map.put("province", father.getProvince());
            map.put("city", father.getCity());
            render = xwpfTemplate.render(map);
            if (niceXWPFDocument == null) {
                niceXWPFDocument = render.getXWPFDocument();
            } else {
                documentList.add(render.getXWPFDocument());
//                niceXWPFDocument = niceXWPFDocument.merge(render.getXWPFDocument());
                render.close();
            }
//            printTime(father.getProvince() + i++);
        }
//        printTime("渲染数据：");
//        niceXWPFDocument = niceXWPFDocument.merge(documentList, niceXWPFDocument.createParagraph().createRun());
        printTime("合并数据：");
        niceXWPFDocument.write(new FileOutputStream("C:\\Users\\13015\\Desktop\\temp\\result.docx"));
        printTime("写入数据：");
    }

    public void test3() throws Exception{
        start = System.currentTimeMillis();
        ZipSecureFile.setMinInflateRatio(0d);
        List<Father> list = initData();
        printTime("生产数据：");
        NiceXWPFDocument niceXWPFDocument = null;
        List<NiceXWPFDocument> documentList = new LinkedList<>();
        XWPFTemplate xwpfTemplate;
        for (Father father : list) {
            xwpfTemplate = XWPFTemplate.compile("C:\\Users\\13015\\Desktop\\temp\\template.docx");
            XWPFTemplate render = xwpfTemplate.render(father);
            if (niceXWPFDocument == null) {
                niceXWPFDocument = render.getXWPFDocument();
            } else {
//                niceXWPFDocument = niceXWPFDocument.merge(render.getXWPFDocument());
                documentList.add(render.getXWPFDocument());
                render.close();
            }
        }
//        printTime("渲染数据：");
        niceXWPFDocument = niceXWPFDocument.merge(documentList, niceXWPFDocument.createParagraph().createRun());
        printTime("合并数据：");
        niceXWPFDocument.write(new FileOutputStream("C:\\Users\\13015\\Desktop\\temp\\result.docx"));
        printTime("写入数据：");
    }

    private List<Father> initData() {
        List<Father> list = new LinkedList<>();
        int i = 0;
        Father father;
        List<Father.Children> childrenList;
        while (i < 60) {
            father = new Father();
            father.setProvince("zhejiang");
            father.setCountry("china");
            father.setCity("hangzhou");
            childrenList = new LinkedList<>();
            for (int j = 0; j < 100; j++) {
                childrenList.add(getChildren());
            }
            father.setChildren(childrenList);
            list.add(father);
            i++;
        }
        return list;
    }

    private Father.Children getChildren() {
        Father.Children children = new Father.Children();
        children.setAge("11");
        children.setName1("name");
        children.setName2("name");
        children.setName3("name");
        children.setName4("name");
        children.setName5("name");
        children.setName6("name");
        children.setName7("name");
        children.setName8("name");
        children.setName9("name");
        return children;
    }

    private void printTime(String message) {
        long end = System.currentTimeMillis();
        System.out.println(message + (end - start) + "ms");
        start = end;
    }


    public static void main(String[] args) {
        long start = System.currentTimeMillis();
        XWPFTemplate xwpfTemplate;
        File file = new File("C:\\Users\\13015\\Desktop\\temp\\template1.docx");
        for (int i = 0; i < 10000; i++) {
            xwpfTemplate = XWPFTemplate.compile(file);
        }
        long end = System.currentTimeMillis();
        System.out.println((end - start) + "ms");
    }
    //21990ms

}
