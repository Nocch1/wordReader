package com.zc.officereader.wordreader;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.lang.String;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 读取word文档中表格数据，支持doc、docx
 *
 * @author Fise19
 */
public class ExportDoc {
    public static void main(String[] args) {
        List<OcrFieldKeyword> keywordList = new ArrayList<>();
        OcrFieldKeyword tem = new OcrFieldKeyword();
        tem.setFieldId(1).setKwValue('1221');
        keywordList.add;


        ExportDoc test = new ExportDoc();
        String home = System.getProperty("user.home");
        String BASE_PATH = home + File.separator + "识别" + File.separator + "可识别" + File.separator;
        String filePath = BASE_PATH + "2.doc";
//		String filePath = "D:\\new\\测试.doc";
        test.testWord(filePath);
    }

    /**
     * 读取文档中表格
     *
     * @param filePath
     */
    public void testWord(String filePath) {
        try {
            FileInputStream in = new FileInputStream(filePath);//载入文档
            // 处理docx格式 即office2007以后版本
            if (filePath.toLowerCase().endsWith("docx")) {
                //word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
                XWPFDocument xwpf = new XWPFDocument(in);//得到word文档的信息
                Iterator<XWPFTable> it = xwpf.getTablesIterator();//得到word中的表格
                // 设置需要读取的表格  set是设置需要读取的第几个表格，total是文件中表格的总数
                int set = 2, total = 4;
                int num = set;
                // 过滤前面不需要的表格
                for (int i = 0; i < set - 1; i++) {
                    it.hasNext();
                    it.next();
                }
                while (it.hasNext()) {
                    XWPFTable table = it.next();
                    System.out.println("这是第" + num + "个表的数据");
                    List<XWPFTableRow> rows = table.getRows();
                    //读取每一行数据
                    for (int i = 0; i < rows.size(); i++) {
                        XWPFTableRow row = rows.get(i);
                        //读取每一列数据
                        List<XWPFTableCell> cells = row.getTableCells();
                        for (int j = 0; j < cells.size(); j++) {
                            XWPFTableCell cell = cells.get(j);
                            //输出当前的单元格的数据
                            System.out.print(cell.getText() + "\t");
                        }
                        System.out.println();
                    }
                    // 过滤多余的表格
                    while (num < total) {
                        it.hasNext();
                        it.next();
                        num += 1;
                    }
                }
            } else {
                // 处理doc格式 即office2003版本
                POIFSFileSystem pfs = new POIFSFileSystem(in);
                HWPFDocument doc = new HWPFDocument(pfs);
                Range range = doc.getRange();//得到文档的读取范围
                TableIterator tableIterator = new TableIterator(range);

                String[][] tableCells = new String[20][];
                // 迭代文档中的表格
                while (tableIterator.hasNext()) {
                    Table tb = (Table) tableIterator.next();
                    for (int i = 0; i < tb.numRows(); i++) {
                        TableRow tr = tb.getRow(i);
                        //迭代列，默认从0开始
                        for (int j = 0; j < tr.numCells(); j++) {
                            TableCell td = tr.getCell(j);//取得单元格
                            String tem = "";
                            //取得单元格的内容
                            for (int k = 0; k < td.numParagraphs(); k++) {
                                Paragraph para = td.getParagraph(k);
                                String s = para.text();
                                //去除后面的特殊符号
                                if (s != null && !"".equals(s)) {
                                    s = s.substring(0, s.length() - 1);
                                }
                                assert s != null;
                                tem = tem.concat(s);
                                System.out.print(s + "\t");
                            }
                            tableCells[i][j] = tem.toLowerCase();
                        }
                        System.out.println();
                    }
                }
                tableHandle(tableCells);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void tableHandle(String[][] arr) {
        Object a;
        for (int i = 0; i < arr.length; i++) {
            for (int j = 0; j < arr[i].length; j++) {

                String txt = arr[i][j];
                String re1 = "(shipper)";    // Word 1

                Pattern p = Pattern.compile(re1, Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
                Matcher m = p.matcher(txt);
                if (m.find()) {
                    String word1 = m.group(1);

                }
            }
        }
    }
}
