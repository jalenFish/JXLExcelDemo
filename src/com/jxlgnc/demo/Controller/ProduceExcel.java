package com.jxlgnc.demo.Controller;

import com.jxlgnc.demo.beans.Physical;
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
    private String uploadPath = "E:\\ouDataTemp"; // 上传文件的目录
    String fileName="";     //默认数据文件名
    String muBanPath="";    //默认模板文件名

    public static void main(String[] args) throws IOException {

    }


    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        uploadXls(request, response);
        createExcel();
    }

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        uploadXls(request, response);
        createExcel();
    }

    // 将一条条excel数据根据模板excel文件生成所有的excel
    public void createExcel(){
        Map<String,Object> map=new HashMap<String,Object>();
        //这个是模板文件
        //String muBanPath = "E:\\000.xls";
        Physical physical=new Physical();

        // 这个是excel数据文件：eg：E:/yujianlin/xxxxx.xls，必须是xls格式
        //如果后缀不是xls，请在excel另存为选择 xls 即可
        File file = new File(uploadPath+"/"+fileName);
        if(file==null||"".equals(file)){
            System.out.println("数据文件不存在，请检查文件是否存在");
            return;
        }
        //obj.readExcel(file);
        InputStream in = null;
        OutputStream out = null;
        try {
            //得到excel数据文件的所有数据
            List<List<String>> allData = GetExcelInfo.readExcel2(file);
            System.out.println(allData);

            for (int i =0;i<allData.size();i++){
                //得到每一条数据
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
                map.put("physical", physical);

                in = ExcelTool.exportExcel(uploadPath+"/"+muBanPath, map);
                out = new FileOutputStream(uploadPath+"/allexcel/"+i+".xls");
                BufferedInputStream bin = new BufferedInputStream(in);
                BufferedOutputStream bout = new BufferedOutputStream(out);
                while (true) {
                    int datum = bin.read();
                    if (datum == -1)
                        break;
                    bout.write(datum);
                }
                //刷新缓冲区
                bout.flush();
            }


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
        //如果目录不存在
        if (!file.isDirectory()) {
            //创建文件上传目录
            file.mkdirs();
        }

        // Create a factory for disk-based file items
        DiskFileItemFactory factory = new DiskFileItemFactory();

        // Set factory constraintsx
        factory.setSizeThreshold(4096); //设置缓冲区大小
        factory.setRepository(tempPathFile);// 设置缓冲区目录

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

            //整个表单的所有域都会被解析，要先判断一下是普通表单域还是文件上传域
            if (item.isFormField()) {
                String name = item.getFieldName();
                String value = item.getString();
                System.out.println(name + ":" + value);
            } else {
                String fieldName = item.getFieldName();
                if ("uploadxlsfile".equals(fieldName)){ //数据文件
                    fileName = item.getName();
                }else if("mubanfile".equals(fieldName)){
                    muBanPath = item.getName();
                }

                String contentType = item.getContentType();
                boolean isInMem = item.isInMemory();
                long sizeInBytes = item.getSize();
                System.out.println(fieldName + ":" + fileName);
                System.out.println("类型：" + contentType);
                System.out.println("是否在内在：" + isInMem);
                System.out.println("文件大小" + sizeInBytes);

                //将文件保存到本地，暂时不用保存到本地
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
