package com.jxlgnc.demo.Controller;

import com.jxlgnc.demo.beans.Physical;
import com.jxlgnc.demo.util.DateTool;
import com.jxlgnc.demo.util.ExcelTool;
import com.jxlgnc.demo.util.GetExcelInfo;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created by yjl on 2018-07-19.
 */
@WebServlet("/ProduceExcel")
public class ProduceExcel extends HttpServlet{
    File tempPathFile;
    private String uploadPath = "E:\\ouDataTemp"; // �ϴ��ļ���Ŀ¼
    String fileName="";     //Ĭ�������ļ���
    String muBanPath="pages/tempFile/000.xls";    //Ĭ��ģ���ļ���
    String filePath=""; //ģ���ļ�Ŀ¼

    public static void main(String[] args) throws IOException {

    }


    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        /*uploadXls(request, response);
        createExcel();
        request.getRequestDispatcher("pages/success.jsp").forward(request,response);*/
        doGet(request,response);
    }

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {

        uploadXls(request, response);
        try {
            createExcel(request,response);
        }catch (Exception e){
            e.printStackTrace();
            request.getRequestDispatcher("pages/error.jsp").forward(request,response);
        }


    }

    // ��һ����excel���ݸ���ģ��excel�ļ��������е�excel
    public void createExcel(HttpServletRequest request,HttpServletResponse response){
        filePath = request.getSession().getServletContext().getRealPath("pages/tempFile/");
        Map<String,Object> map=new HashMap<String,Object>();
        //�����ģ���ļ�
        //String muBanPath = "E:\\000.xls";
        Physical physical=new Physical();

        // �����excel�����ļ���eg��E:/yujianlin/xxxxx.xls��������xls��ʽ
        //�����׺����xls������excel���Ϊѡ�� xls ����
        File file = new File(uploadPath+"/"+fileName);
        if(file==null||"".equals(file)){
            System.out.println("The file does not exist!please check again.");
            return;
        }
        //obj.readExcel(file);
        InputStream in = null;
        OutputStream out = null;
        try {
            //�õ�excel�����ļ�����������
            List<List<String>> allData = GetExcelInfo.readExcel2(file);
            System.out.println(allData);

            for (int i =0;i<allData.size();i++){
                //�õ�ÿһ������
                List<String> eachData = allData.get(i);

                physical.setA1(eachData.get(0));
                physical.setA2(eachData.get(1));
                physical.setA3(eachData.get(2));
                physical.setA4(eachData.get(3));
                physical.setA5(eachData.get(4));
                physical.setA6(eachData.get(5));
                physical.setA7(eachData.get(6));
                physical.setA8(eachData.get(7));
                physical.setA9(eachData.get(8));
                physical.setA10(eachData.get(9));
                physical.setA11(DateTool.getTodayDate().toString());
                map.put("physical", physical);

                //in = ExcelTool.exportExcel(uploadPath+"/"+muBanPath, map);
                in = ExcelTool.exportExcel( filePath+"/000.xls", map);
                //CPR-SEASON����һ�У�-ESQ-VENDOR����8��-BOARD / MODEL����2��
                out = new FileOutputStream(uploadPath+"/allexcel/"+"CPR-"+eachData.get(0)+"-ESQ-"+eachData.get(7)+"-"+eachData.get(1)+".xls"); //����������
                BufferedInputStream bin = new BufferedInputStream(in);
                BufferedOutputStream bout = new BufferedOutputStream(out);
                while (true) {
                    int datum = bin.read();
                    if (datum == -1)
                        break;
                    bout.write(datum);
                }
                //ˢ�»�����
                bout.flush();
            }

            request.getRequestDispatcher("pages/success.jsp").forward(request,response);
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();

        }finally {
            if (in != null) {

                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

        }

    }

    public void uploadXls(HttpServletRequest request,
                          HttpServletResponse response) {
        boolean isMultipart = ServletFileUpload.isMultipartContent(request);
        if (!isMultipart) {
            return;
        }
        File file = new File(uploadPath+"/allexcel");
        //���Ŀ¼������
        if (!file.isDirectory()) {
            //�����ļ��ϴ�Ŀ¼
            file.mkdirs();
        }

        // Create a factory for disk-based file items
        DiskFileItemFactory factory = new DiskFileItemFactory();

        // Set factory constraintsx
        factory.setSizeThreshold(4096); //���û�������С
        factory.setRepository(tempPathFile);// ���û�����Ŀ¼

        // Create a new file upload handler
        ServletFileUpload upload = new ServletFileUpload(factory);


        // Parse the request
        List<FileItem> items = null;
        try {
            items = upload.parseRequest(request);
        } catch (FileUploadException e1) {
            e1.printStackTrace();
        }

        // Process the uploaded items
        Iterator<FileItem> iter = items.iterator();
        while (iter.hasNext()) {
            FileItem item = (FileItem)iter.next();

            //�������������򶼻ᱻ������Ҫ���ж�һ������ͨ�������ļ��ϴ���
            if (item.isFormField()) {
                String name = item.getFieldName();
                String value = item.getString();
                System.out.println(name + ":" + value);
            } else {
                String fieldName = item.getFieldName();
                if ("uploadxlsfile".equals(fieldName)){ //�����ļ�
                    fileName = item.getName();
                }else if("mubanfile".equals(fieldName)){
                    muBanPath = item.getName();
                }

                String contentType = item.getContentType();
                boolean isInMem = item.isInMemory();
                long sizeInBytes = item.getSize();
                System.out.println(fieldName + ":" + fileName);
                System.out.println("���ͣ�" + contentType);
                System.out.println("�Ƿ������ڣ�" + isInMem);
                System.out.println("�ļ���С" + sizeInBytes);

                //���ļ����浽���أ���ʱ���ñ��浽����
                try {
                    byte[] bytes = item.getName().getBytes();
                    File fullfile = new File(new String(bytes, "utf-8"));
                    File uploadfile = new File(uploadPath,fullfile.getName());
                    item.write(uploadfile);

                }catch (Exception e){
                    e.printStackTrace();
                }

            }
        }
    }
}
