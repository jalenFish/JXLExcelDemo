package com.jxlgnc.demo.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

//import com.hna.aircrewhealth.po.AviatorHealthCheck;
//import com.hna.aircrewhealth.security.po.Staff;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCell;

import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class ExcelTool {

    private static String path = "F:\\luokq\\aircrewhealth\\template\\�ش󼲲����浥.xls";
    private final static String defaultName = "��ҽ�������¼��.xls";

    public static String getPath() {
        return path;
    }

    public static void setPath(String p) {
        path = p;
    }
//���MAIN�����Ǹ�DEMO ���Բ������д�� 
//    public static void main(String[] arg) {
//        AviatorHealthCheck bean = new AviatorHealthCheck();
//
//        bean.setId("3FCB19B440E74DF1BD50CD123A3C087C");
//        bean.setHealthCheckFlightNum("NB-38-54321");
//
//        Staff s = new Staff(); //����һ��ʵ����
//        s.setName("����");     
//        bean.setFollowDoctor(s);
//
//        bean.setNoddeid("878787");
//        bean.setStartTime("1987-02-25");
//        bean.setEndTime("1987-02-25");
//        bean.setFollowContext("���ɴ�������˯�����������¶��ƿȿȿ�ʲô�̼������Ŀ�˹�Ƶ����羰�Ŀ�˹�ƣ����ɷ����");
//
//        Map<String, Object> map = new HashMap<String, Object>();   //����һ��map
//        map.put("AviatorHealthCheck", bean);  //��ʵ�����Map,��Ϊ�ǰ�Map�����
/**
*���������ѭ���õ�
*/
//        List<Object> list = new ArrayList<Object>();
//
//        for (int i = 0; i < 6; i++) {
//            bean = new AviatorHealthCheck();
//
//            bean.setId("3FCB19B440E74DF1BD50CD123A3C087C"+"----"+i);
//            bean.setHealthCheckFlightNum("NB-38-54321"+"----"+i);
//
//            s = new Staff();
//            s.setName("����"+"----"+i);
//            bean.setFollowDoctor(s);
//
//            bean.setNoddeid("878787"+"----"+i);
//            bean.setStartTime("1987-02-25"+"----"+i);
//            bean.setEndTime("1987-02-25"+"----"+i);
//            bean.setFollowContext("���ɴ�������˯�����������¶��ƿȿȿ�ʲô�̼������Ŀ�˹�Ƶ����羰�Ŀ�˹�ƣ����ɷ����"+"----"+i);
//            list.add(bean);
//        }
//        map.put("listname", list);
//        map.put("listname2", list);
/**����ѭ��������*/
//        exportExcel(path, map);   //�������Ҫ���ǵ���Excel�ķ���
//    }

    /**
     * ���� Excel
     * 
     * @param template
     *            Excelģ��
     * @param datas
     *            ����
     * @return
     */
    public static FileInputStream exportExcel(String template, Map<String, Object> datas) {
        FileInputStream fis = null;
        InputStream is = FileTool.getFileInputStream(template);

        try {
            if (is != null) {
                Workbook book = Workbook.getWorkbook(is);
                File tempFile = File.createTempFile("temp", ".xls");
                
                WritableWorkbook wWorkbook = Workbook.createWorkbook(tempFile, book);

                /** �������ʽ�����͵����ݡ� **/
                generateExpData(book, wWorkbook, datas);
                /** ����ѭ������������͵����ݡ� **/
                generateEachData(book, wWorkbook, datas);

                wWorkbook.write();
                wWorkbook.close();
            
                fis = new FileInputStream(tempFile);
            }
        } catch (BiffException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (WriteException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } finally{
            if(is!=null){
                try {
                    is.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }

        return fis;
    }

    public static FileInputStream exportExcel1(String template, Map<String, Object> datas) {
        FileInputStream fis = null;
        InputStream is = FileTool.getFileInputStream(template);

        try {
            if (is != null) {
                Workbook book = Workbook.getWorkbook(is);
                File tempFile = File.createTempFile("temp", ".xls");
                WritableWorkbook wWorkbook = Workbook.createWorkbook(tempFile, book);

                /** �������ʽ�����͵����ݡ� **/
                generateExpData1(book, wWorkbook, datas);
                /** ����ѭ������������͵����ݡ� **/
                generateEachData(book, wWorkbook, datas);

                wWorkbook.write();
                wWorkbook.close();

                fis = new FileInputStream(tempFile);
            }
        } catch (BiffException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (WriteException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } finally{
            if(is!=null){
                try {
                    is.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }

        return fis;
    }
    /**
     * �������ʽ�����͵����ݡ�
     * 
     * @param book
     *            ��ģ�塿����
     * @param wWorkbook
     *            ����ģ�崴���ġ��������ļ�������
     */
    private static void generateExpData(Workbook book, WritableWorkbook wWorkbook, Map<String, Object> datas) throws Exception {
        List<ExcelCells> expcells = search("${", book);
        for (ExcelCells cell : expcells) {
            wWorkbook.getSheet(cell.getSheetIndex()).addCell(getValueByExp(cell, datas));
        }
    }
    
    private static void generateExpData1(Workbook book, WritableWorkbook wWorkbook, Map<String, Object> datas) throws Exception {
        List<ExcelCells> expcells = search("${", book);
        for (ExcelCells cell : expcells) {
            wWorkbook.getSheet(cell.getSheetIndex()).addCell(getValueByExp1(cell, datas));
        }
    }

    /**
     * ����ѭ������������͵�����
     * 
     * @param book
     *            ��ģ�塿����
     * @param wWorkbook
     *            ����ģ�崴���ġ��������ļ�������
     */
    private static void generateEachData(Workbook book, WritableWorkbook wWorkbook, Map<String, Object> datas) throws Exception {
        List<ExcelCells> each = search("each.", book);
        /* �ȶ�ģ���ж��󣬽��з��顣 */
        Map<String, List<ExcelCells>> map = new LinkedHashMap<String, List<ExcelCells>>();//
        for (ExcelCells cell : each) {
            String[] array = cell.getCell().getContents().trim().split("\\.");
            if (array.length >= 3) {
                List<ExcelCells> list = map.get(array[0] + "." + array[1]);
                if (list == null) {
                    list = new ArrayList<ExcelCells>();
                    map.put(array[0] + "." + array[1], list);
                }
                list.add(cell);
            }
        }

        Iterator<String> iterator = map.keySet().iterator();

        int insertrow = 0;//��ʶ��ǰ�����������˶��������ݡ�
        int lastSheetIndex = -1;//��ʶ��һ�ι�������±ꡣ

        while (iterator.hasNext()) {
            List<ExcelCells> list = map.get(iterator.next());
            int sheetIndex = list.get(0).getSheetIndex();// ��ȡ����±ꡣ
            //���л���������  insertrow �� 0
            if(lastSheetIndex != -1 && lastSheetIndex != sheetIndex) insertrow = 0;
            lastSheetIndex = sheetIndex;
            
            int startRow = list.get(0).getCell().getRow() + insertrow;// ��ȡ��ʼ���±ꡣ

            String[] array = list.get(0).getCell().getContents().trim().split("\\.");
            if (array.length > 0) {
                Object data = datas.get(array[1]);
                if (data != null && !data.getClass().getName().equals(List.class.getName()) && !data.getClass().getName().equals(ArrayList.class.getName())) {
                    throw new Exception("���ݣ�" + array[1] + "����һ�������࣡");
                }
                List<Object> rowsData = (List<Object>) data;
                // ������ʱ��
                if (rowsData != null && rowsData.size() > 0) {
                    for (int i = 0; i < rowsData.size(); i++) {
                        /* ��һ�����ݣ�����ģ��λ�ã����Բ���Ҫ�������� */
                        if (i == 0) {
                            for (ExcelCells cell : list) {
                                wWorkbook.getSheet(sheetIndex).addCell(getValueByEach(cell, rowsData.get(i), startRow, cell.getCell().getColumn()));
                            }
                            continue;
                        }
                        /* �������� */
                        wWorkbook.getSheet(sheetIndex).insertRow(startRow + i);
                        for (ExcelCells cell : list) {
                            wWorkbook.getSheet(sheetIndex).addCell(getValueByEach(cell, rowsData.get(i), startRow + i, cell.getCell().getColumn()));
                        }
                        insertrow++;
                    }

                }
                // ������ʱ��
                else {
                    for (ExcelCells cell : list) {
                        wWorkbook.getSheet(sheetIndex).addCell(getValueByEach(cell, null, startRow, cell.getCell().getColumn()));
                    }
                }
            }
        }

    }

    /**
     * ���ݡ����ʽ�������ݼ��л�ȡ��Ӧ���ݡ�
     * 
     * @param exp
     *            ���ʽ
     * @param datas
     *            ���ݼ�
     * @return
     */
    public static WritableCell getValueByExp(ExcelCells cells, Map<String, Object> datas) {
        WritableCell writableCell = null;

        List<Object> values = new ArrayList<Object>();
        List<String> exps = cells.getExps();// ��ȡ���ʽ���ϡ�

        String old_c = cells.getCell().getContents();// ģ��ԭ���ݡ�

        System.out.println(old_c);
        
        for (String exp : exps) {

            String[] names = exp.replace("${", "").replace("}", "").split("\\.");

            Object object = null;
            for (String name : names) {
                if (object == null){
                    object = ObjectCustomUtil.getValueByFieldName(name, datas);
                    System.out.println(object);
                }else{
                    object = ObjectCustomUtil.getValueByFieldName(name, object);
                }
            
            }
            // ${asd.sdfa}
            if (!old_c.isEmpty()) {
                while (old_c.indexOf(exp) != -1)
                    old_c = old_c.replace(exp, object.toString());
            }
        }

        writableCell = getWritableCellByObject(cells.getCell().getRow(), cells.getCell().getColumn(), old_c);
        writableCell.setCellFormat(cells.getCell().getCellFormat());

        return writableCell;
    }
    
    /*
     * ���������ר������פ����黷���������
     */
    public static WritableCell getValueByExp1(ExcelCells cells, Map<String, Object> datas) {
        WritableCell writableCell = null;

        List<Object> values = new ArrayList<Object>();
        List<String> exps = cells.getExps();// ��ȡ���ʽ���ϡ�

        String old_c = cells.getCell().getContents();// ģ��ԭ���ݡ�

        for (String exp : exps) {

            String[] names = exp.replace("${", "").replace("}", "").split("\\.");

            Object object = null;
            String checkContentValue = "";
            for (String name : names) {
                if (object == null){
                    object = ObjectCustomUtil.getValueByFieldName(name, datas);
                }else{
                    object = ObjectCustomUtil.getValueByFieldName(name, object);
                }
                if(name.indexOf("checkContent")!=-1){
                    if("0".equals(object.toString())){
                        checkContentValue = "����";
                    }else if("1".equals(object.toString())){
                        checkContentValue = "������";
                    }else{
                        checkContentValue = "δ���";
                    }
                }else if(name.indexOf("checkTime")!=-1){
                     Date date = (Date)object;
                     checkContentValue = date.getYear()+"��"+ (date.getMonth()+1) +"��" +date.getDate();
                }
            }
            
            if (!old_c.isEmpty()) {
                while (old_c.indexOf(exp) != -1){
                    if("".equals(checkContentValue)){
                        old_c = old_c.replace(exp, object.toString());
                    }else{
                        old_c = old_c.replace(exp, checkContentValue);
                    }
                }
            }
        }

        writableCell = getWritableCellByObject(cells.getCell().getRow(), cells.getCell().getColumn(), old_c);
        writableCell.setCellFormat(cells.getCell().getCellFormat());

        return writableCell;
    }

    /**
     * ���ݡ�Each���ʽ�������ݼ��л�ȡ��Ӧ���ݡ�
     * 
     * @param exp
     *            ���ʽ
     * @param datas
     *            ���ݼ�
     * @return
     */
    public static WritableCell getValueByEach(ExcelCells cells, Object datas, int rows, int column) {
        WritableCell writableCell = null;

        if (datas != null) {
            List<Object> values = new ArrayList<Object>();
            String[] exps = cells.getCell().getContents().trim().split("\\.");// ��ȡ���ʽ���ϡ�

            Object object = null;
            for (int i = 2; i < exps.length; i++) {
                if (object == null)
                    object = ObjectCustomUtil.getValueByFieldName(exps[i], datas);
                else
                    object = ObjectCustomUtil.getValueByFieldName(exps[i], object);
            }
            writableCell = getWritableCellByObject(rows, column, object);
        } else {
            writableCell = getWritableCellByObject(rows, column, null);
        }
        writableCell.setCellFormat(cells.getCell().getCellFormat());

        return writableCell;
    }

    /**
     * ��δʵ�֡�
     * 
     * @param beginRow
     * @param beginColumn
     * @param heads
     * @param result
     * @return
     */
    public static synchronized String customExportExcel(int beginRow, int beginColumn, Map heads, List result) {
        return null;
    }

    /**
     * �����ṩ�ġ��б꡿�����б꡿��������ֵ������һ��Excel�ж���
     * 
     * @param beginRow
     *            ���б꡿
     * @param beginColumn
     *            ���б꡿
     * @param obj
     *            ������ֵ��
     * @return
     */
    public static WritableCell getWritableCellByObject(int beginRow, int beginColumn, Object obj) {
        WritableCell cell = null;

        if (obj == null)
            return new Label(beginColumn, beginRow, "");
        if (obj.getClass().getName().equals(String.class.getName())) {
            cell = new Label(beginColumn, beginRow, obj.toString());
        } else if (obj.getClass().getName().equals(int.class.getName()) || obj.getClass().getName().equals(Integer.class.getName())) {
            // jxl.write.Number
            cell = new Number(beginColumn, beginRow, Integer.parseInt(obj.toString()));
        } else if (obj.getClass().getName().equals(float.class.getName()) || obj.getClass().getName().equals(Float.class.getName())) {
            cell = new Number(beginColumn, beginRow, Float.parseFloat(obj.toString()));
        } else if (obj.getClass().getName().equals(double.class.getName()) || obj.getClass().getName().equals(Double.class.getName())) {
            cell = new Number(beginColumn, beginRow, Double.parseDouble(obj.toString()));
        } else if (obj.getClass().getName().equals(long.class.getName()) || obj.getClass().getName().equals(Long.class.getName())) {
            cell = new Number(beginColumn, beginRow, Long.parseLong(obj.toString()));
        } else if (obj.getClass().getName().equals(Date.class.getName())) {
            cell = new DateTime(beginColumn, beginRow, (Date)obj);
        } else {
            cell = new Label(beginColumn, beginRow, obj.toString());
        }
        return cell;
    }

    /**
     * ����ĳ�ַ���һ�γ��ֵ�λ�á�
     * 
     * @param text
     *            ���ı���
     * @param book
     *            ��Excel����
     * @return
     */
    public static ExcelCells searchFirstText(String text, Workbook book) {
        ExcelCells Rcell = null;
        Sheet[] sheets = book.getSheets();
        if (sheets != null) {
            int sheetIndex = 0;
            for (Sheet sheet : sheets) {
                if (sheet != null) {
                    int rows = sheet.getRows();
                    if (rows > 0) {
                        for (int i = 0; i < rows; i++) {
                            Cell[] cells = sheet.getRow(i);
                            if (cells != null) {
                                for (Cell cell : cells) {
                                    if (cell != null && !StringUtils.isNull(cell.getContents())) {
                                        String contents = cell.getContents();
                                        if (contents.equals(text))
                                            return new ExcelCells(sheet, cell, sheetIndex);
                                    }
                                }
                            }
                        }
                    }
                }
                sheetIndex++;
            }
        }
        return Rcell;
    }

    /**
     * ���Ұ���ĳ�ַ����е��ж���
     * 
     * @param text
     *            ���ı���
     * @param book
     *            ��Excel����
     * @return
     */
    public static List<ExcelCells> search(String text, Workbook book) {
        List<ExcelCells> rcells = new ArrayList<ExcelCells>();

        Sheet[] sheets = book.getSheets();
        if (sheets != null)
            for (Sheet sheet : sheets) {
                if (sheet != null) {
                    int rows = sheet.getRows();
                    if (rows > 0) {
                        for (int i = 0; i < rows; i++) {
                            Cell[] cells = sheet.getRow(i);
                            if (cells != null) {
                                for (Cell cell : cells) {
                                    if (cell != null && !StringUtils.isNull(cell.getContents())) {
                                        String contents = cell.getContents();
                                        if (contents.indexOf(text) != -1)
                                            rcells.add(new ExcelCells(sheet, cell));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        return rcells;
    }

}
