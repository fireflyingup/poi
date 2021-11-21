package com.poi.study.service;


import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.stereotype.Service;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.LinkedList;
import java.util.List;

@Service
public class PoiService{

    private static long start = 0L;

    public void test() throws Exception {
        XWPFDocument xwpfDocument = new XWPFDocument();
        XWPFStyles styles = xwpfDocument.createStyles();
        CTStyles ctStyles = CTStyles.Factory.newInstance();
        styles.setStyles(ctStyles);

        CTBody body = xwpfDocument.getDocument().getBody();
        CTSectPr ctSectPr = body.addNewSectPr();
        ctSectPr.addNewPgSz();
        CTPageSz pgSz = ctSectPr.getPgSz();
        // 设置页面大小为A4纸
        pgSz.setW(BigInteger.valueOf(11907));
        pgSz.setH(BigInteger.valueOf(16840));
        pgSz.setOrient(STPageOrientation.PORTRAIT);

        if (!ctSectPr.isSetPgMar()) {
            ctSectPr.addNewPgMar();
        }
        // 设置页面边距
        CTPageMar pgMar = ctSectPr.getPgMar();
        pgMar.setLeft(BigInteger.valueOf(720L));
        pgMar.setTop(BigInteger.valueOf(1440L));
        pgMar.setRight(BigInteger.valueOf(720L));
        pgMar.setBottom(BigInteger.valueOf(1440L));


        XWPFTable table = xwpfDocument.createTable(4, 2);
        table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(8600 / 2));
//        //other columns (2 in this case) also each 1 inches width
        for (int col = 1; col < 2; col++) {
            table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(8600 / 2));
        }
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            List<XWPFTableCell> tableCells = row.getTableCells();
            for (int j = 0; j < tableCells.size(); j++) {
                XWPFTableCell cell = tableCells.get(j);
                CTTcPr ctTcPr = cell.getCTTc().addNewTcPr();
                if (i == 0 || i == 2) {
//                    if (ctTcPr.getGridSpan() == null) {
//                        CTDecimalNumber ctDecimalNumber = CTDecimalNumber.Factory.newInstance();
//                        ctDecimalNumber.setVal(BigInteger.valueOf(1));
//                        ctTcPr.setGridSpan(ctDecimalNumber);
//                    } else {
//                        ctTcPr.getGridSpan().setVal(BigInteger.valueOf(1));
//                    }
                    continue;
                }
                if (j == 0) {
                    CTDecimalNumber ctDecimalNumber = ctTcPr.getGridSpan();
                    if (!ctTcPr.isSetGridSpan()) {
                        ctDecimalNumber = CTDecimalNumber.Factory.newInstance();
                    }
                    ctDecimalNumber.setVal(BigInteger.valueOf(2));
                    ctTcPr.setGridSpan(ctDecimalNumber);
                    XWPFParagraph xwpfParagraph = cell.getParagraphs().get(0);
                    XWPFRun run = xwpfParagraph.createRun();
                    run.setText("asdsadasd");
                    run.setFontFamily("黑体");
                } else {
                    // 移除调没有的列
//                    ctDecimalNumber.setVal(BigInteger.valueOf(0));
//                    ctTcPr.setGridSpan(ctDecimalNumber);
                    row.getCtRow().removeTc(j);
                    row.removeCell(j);
                }
            }
        }

        FileOutputStream fos = new FileOutputStream("/opt/astp/data/report/generate/a.docx");
        xwpfDocument.write(fos);

        FileOutputStream fos1 = new FileOutputStream("/opt/astp/data/report/generate/a.pdf");

        PdfOptions pdfOptions = PdfOptions.create();

        pdfOptions.fontProvider(new IFontProvider() {
            @Override
            public Font getFont(String s, String s1, float v, int i, Color color) {
                try {
                    if (s.equalsIgnoreCase("宋体") || s.equalsIgnoreCase("微软雅黑")) {
                        BaseFont baseFont =
                                BaseFont.createFont("/opt/astp/data/report/font/SIMSUN.TTC,0", s1, true);
                        return new Font(baseFont, v, i, color);
                    }
                    if (s.equalsIgnoreCase("黑体")) {
                        BaseFont baseFont =
                                BaseFont.createFont("/opt/astp/data/report/font/SIMHEI.TTF", s1, true);
                        return new Font(baseFont, v, i, color);
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
                if (s == null) {
                    s = "DengXian";
                }
                return FontFactory.getFont(s, s1, v, i, color);
            }
        });
        PdfConverter.getInstance().convert(xwpfDocument, fos1, pdfOptions);
    }

    public void test2() throws Exception{
        XWPFDocument xwpfDocument = new XWPFDocument(new FileInputStream("/opt/astp/data/report/generate/aa.docx"));

        FileOutputStream fos1 = new FileOutputStream("/opt/astp/data/report/generate/aa.pdf");

        PdfOptions pdfOptions = PdfOptions.create();

        pdfOptions.fontProvider(new IFontProvider() {
            @Override
            public Font getFont(String s, String s1, float v, int i, Color color) {
                try {
                    if (s.equalsIgnoreCase("宋体") || s.equalsIgnoreCase("微软雅黑")) {
                        BaseFont baseFont =
                                BaseFont.createFont("/opt/astp/data/report/font/SIMSUN.TTC,0", s1, true);
                        return new Font(baseFont, v, i, color);
                    }
                    if (s.equalsIgnoreCase("黑体")) {
                        BaseFont baseFont =
                                BaseFont.createFont("/opt/astp/data/report/font/SIMHEI.TTF", s1, true);
                        return new Font(baseFont, v, i, color);
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
                if (s == null) {
                    s = "DengXian";
                }
                return FontFactory.getFont(s, s1, v, i, color);
            }
        });
        PdfConverter.getInstance().convert(xwpfDocument, fos1, pdfOptions);
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



}
