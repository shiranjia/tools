package com.tools.doc;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Created by jiashiran on 2016/11/16.
 */
public class DocWriter {

    private static final Set<String> R_C = new HashSet<String>();//入参对象
    private static final String[] inHeader = {"字段","类型","描述","可为空","备注"};//入参对象header
    private static final String[] outHeader = {"字段","类型","描述","备注"};//出参对象header
    static {//入参对象区分表格头属性
        R_C.add("SearchParameter");
        R_C.add("OrderMsg");
        R_C.add("Customer");
        R_C.add("TicketUser");
        R_C.add("TicketUserAddress");
        R_C.add("Invoice");
        R_C.add("CartItem");
        R_C.add("TicketCoupon");
        R_C.add("TicketOrderQuery");

    }

    public static void main(String[] args) {
        createDoc("E:\\git_source\\projectName","docName");
    }

    private static void createDoc(String path, String docName){
        XWPFDocument doc = new XWPFDocument();
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(docName + ".docx");
            File file = new File(path);
            analysisFile(doc,file);
            doc.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if(out!=null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    private static void analysisFile(XWPFDocument doc,File path){
        if(path.isFile()){
            Table table = analysisClass(path);
            String desc = table.getClassName()+"属性说明如下：";
            if(R_C.contains(table.getClassName())){//in
                try {
                    writeDocTable(desc,doc,inHeader,table.getFields());
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }else {//out
                try {
                    writeDocTable(desc,doc,outHeader,table.getFields());
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }else {
            File[] list = path.listFiles();
            for(File f:list){
                analysisFile(doc,f);
            }
        }

    }


    private static Table analysisClass(File file){
        Table table = new Table();
        List<String[]> fields = new ArrayList<String[]>();
        try {
            FileInputStream in = new FileInputStream(file);
            byte[] bytes = new byte[1024 * 1024];
            int size = in.read(bytes);
            String code = new String(bytes,0,size,"gbk");
            String[] divs = code.split("/\\*\\*");
            int i = 1;
            String className = "";
            for(;i<divs.length;i++){
                String div = divs[i];
                if(i == 1){
                    className = analysisClassName(div);
                    className = className.trim().replaceAll(" ","");
                    table.setClassName(className);
                    System.out.println(className);
                }else {
                    String[] field = analysisField(div,className);
                    fields.add(field);
                    for(String f:field){
                        System.out.print(f+"\t");
                    }
                    System.out.println();
                }
            }
            table.setFields(fields);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return table;
    }

    private static String analysisClassName(String line){
        int c = line.indexOf("class") + 6;
        line = line.substring(c,line.length());
        c = line.indexOf(" ");
        line = line.substring(0,c);
        //System.out.println("className="+line.trim());
        return line.trim();
    }

    private static String[] analysisField(String line, String className){
        String[] field = null;
        try {
            int _s = line.indexOf("*");
            int _e = line.indexOf(";");
            line = line.substring(_s, _e);
            String[] lines = line.split("\\n");
            int length = lines.length;
            String fieldBeizhu ="";
            if (length > 3) {
                String beizhu = lines[1].replace("\r","").replace("*","");
                //System.out.println(beizhu);
                fieldBeizhu = beizhu;
            }
            String fieldMiaoshu = lines[0].substring(1, lines[0].length());
            fieldMiaoshu = fieldMiaoshu.replace("\r","");
            String fieldNutNull = "";
            if(fieldMiaoshu.contains("NOT NULL")){
                fieldNutNull = "NO";
                fieldMiaoshu = fieldMiaoshu.replace("NOT NULL","");
            }
            //System.out.println(miaoshu);
            //System.out.println(lines[length-1]);
            String t = lines[length - 1];
            t = t.substring(t.indexOf("p"),t.length());
            String[] f = t.split(" ");
            String fieldType = f[1];
            String fieldName = f[2];
            //System.out.println(type);
            if(R_C.contains(className)){//入参
                field = new String[5];
                field[0] = fieldName;
                field[1] = fieldType;
                field[2] = fieldMiaoshu;
                field[3] = fieldNutNull;
                field[4] = fieldBeizhu;
            }else {
                field = new String[4];
                field[0] = fieldName;
                field[1] = fieldType;
                field[2] = fieldMiaoshu;
                field[3] = fieldBeizhu;
            }
        }catch (Exception e){
            e.printStackTrace();
            System.out.println("exception line = "+line);
        }

        return field;
    }


    /**
     * 生成表格
     * @param desc  描述
     * @param doc   文档
     * @param header
     * @param tables    数据
     * @throws Exception
     */
    public static void writeDocTable(String desc, XWPFDocument doc, String[] header, List<String[]> tables) throws Exception {
        XWPFParagraph tp = doc.createParagraph();
        XWPFRun tRun = tp.createRun();
        tRun.setText(desc);
        tRun.addTab();
        int nRows = tables.size() + 1;//valueSize + a row of header
        int nCols = header.length;
        XWPFTable table = doc.createTable(nRows, nCols);

        // Set the table style. If the style is not defined, the table style
        // will become "Normal".
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        CTString styleStr = tblPr.addNewTblStyle();
        styleStr.setVal("StyledTable");

        // Get a list of the rows in the table
        List<XWPFTableRow> rows = table.getRows();
        int rowCt = 0;
        int colCt = 0;
        for (XWPFTableRow row : rows) {
            // get table row properties (trPr)
            CTTrPr trPr = row.getCtRow().addNewTrPr();
            // set row height; units = twentieth of a point, 360 = 0.25"
            CTHeight ht = trPr.addNewTrHeight();
            ht.setVal(BigInteger.valueOf(360));

            // get the cells in this row
            List<XWPFTableCell> cells = row.getTableCells();
            // add content to each cell
            for (XWPFTableCell cell : cells) {
                // get a table cell properties element (tcPr)
                CTTcPr tcpr = cell.getCTTc().addNewTcPr();
                // set vertical alignment to "center"
                CTVerticalJc va = tcpr.addNewVAlign();
                va.setVal(STVerticalJc.CENTER);
                // create cell color element
                CTShd ctshd = tcpr.addNewShd();
                ctshd.setColor("auto");
                ctshd.setVal(STShd.CLEAR);
                if (rowCt == 0) {
                    // header row
                    ctshd.setFill("A7BFDE");
                }
                else if (rowCt % 2 == 0) {
                    // even row
                    ctshd.setFill("D3DFEE");
                }
                else {
                    // odd row
                    ctshd.setFill("EDF2F8");
                }

                // get 1st paragraph in cell's paragraph list
                XWPFParagraph para = cell.getParagraphs().get(0);
                // create a run to contain the content
                XWPFRun rh = para.createRun();
                // style cell as desired
                if (colCt == nCols - 1) {
                    // last column is 10pt Courier
                    rh.setFontSize(10);
                    rh.setFontFamily("Courier");
                }
                if (rowCt == 0) {
                    //set header row
                    String head = header[colCt];
                    rh.setText(new String(head.getBytes("utf-8"),"utf-8"));
                    rh.setBold(true);
                    para.setAlignment(ParagraphAlignment.CENTER);
                }
                else if (rowCt % 2 == 0) {
                    // even row
                    //rh.setText("row " + rowCt + ", col " + colCt);
                    para.setAlignment(ParagraphAlignment.LEFT);
                }
                else {
                    // odd row
                    //rh.setText("row " + rowCt + ", col " + colCt);
                    para.setAlignment(ParagraphAlignment.LEFT);
                }
                if(rowCt != 0){//set body value
                    String[] values = tables.get(rowCt - 1);
                    String value = values[colCt];
                    if (value != null) {
                        rh.setText(new String(value.getBytes("utf-8"), "utf-8"));
                    }
                }
                colCt++;
            } // for cell
            colCt = 0;
            rowCt++;
        } // for row
    }

    static class Table {

        private String className;
        private List<String[]> fields;

    /*public Table(String name,String[] va){
        this.className = name;
        this.fields = va;
    }*/

        public String getClassName() {
            return className;
        }

        public void setClassName(String className) {
            this.className = className;
        }

        public List<String[]> getFields() {
            return fields;
        }

        public void setFields(List<String[]> fields) {
            this.fields = fields;
        }
    }
}
