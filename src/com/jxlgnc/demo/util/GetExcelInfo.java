package com.jxlgnc.demo.util;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

/**
 * Created by yjl on 2018-07-18.
 */
public class GetExcelInfo {
    /*public static void main(String[] args) {
        GetExcelInfo obj = new GetExcelInfo();
        // �����excel�����ļ���E:/zhanhj/studysrc/jxl��
        File file = new File("E:/111.xls");
        //obj.readExcel(file);
        try {
            //�õ���������
            List<List<String>> allData=readExcel2(file);
            System.out.println(allData);

        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }*/
    // ȥ��Excel�ķ���readExcel���÷�������ڲ���Ϊһ��File����
    public static void readExcel(File file) {
        try {
            // ��������������ȡExcel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            // jxl�ṩ��Workbook��
            Workbook wb = Workbook.getWorkbook(is);
            // Excel��ҳǩ����
            int sheet_size = wb.getNumberOfSheets();
            for (int index = 0; index < sheet_size; index++) {
                // ÿ��ҳǩ����һ��Sheet����
                Sheet sheet = wb.getSheet(index);
                // sheet.getRows()���ظ�ҳ��������
                for (int i = 0; i < sheet.getRows(); i++) {
                    // sheet.getColumns()���ظ�ҳ��������
                    for (int j = 0; j < sheet.getColumns(); j++) {
                        String cellinfo = sheet.getCell(j, i).getContents();
                        System.out.println(cellinfo);
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }



    /**
     * ��ȡ����
     * @param file
     * @return
     * @throws Exception
     */
    public static List<List<String>> readExcel2(File file) throws Exception {

        // ��������������ȡExcel
        InputStream is = new FileInputStream(file.getAbsolutePath());
        // jxl�ṩ��Workbook��
        //Workbook wb = Workbook.getWorkbook(is);
        // ��ȡexcel�ļ���Ԫ��Ӣ�ĳ�����������,���Ի�������������
        WorkbookSettings workbookSettings = new WorkbookSettings();
        workbookSettings.setEncoding("ISO-8859-1");
        Workbook wb = Workbook.getWorkbook(is,workbookSettings);

        // ֻ��һ��sheet,ֱ�Ӵ���
        //����һ��Sheet����
        Sheet sheet = wb.getSheet(0);
        // �õ����е�����
        int rows = sheet.getRows();
        // ���е�����
        List<List<String>> allData = new ArrayList<List<String>>();
        // Խ����һ�� ����������
        for (int j = 1; j < rows; j++) {

            List<String> oneData = new ArrayList<String>();
            // �õ�ÿһ�еĵ�Ԫ�������
            Cell[] cells = sheet.getRow(j);
            for (int k = 0; k < cells.length; k++) {

                oneData.add(cells[k].getContents().trim());
            }
            // �洢ÿһ������
            allData.add(oneData);
            // ��ӡ��ÿһ������
            //System.out.println(oneData);

        }
        return allData;

    }
}
