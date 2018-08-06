package com.jxlgnc.demo.gui;
import com.jxlgnc.demo.beans.Physical;
import com.jxlgnc.demo.util.DateTool;
import com.jxlgnc.demo.util.ExcelTool;
import com.jxlgnc.demo.util.GetExcelInfo;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.*;

import javax.swing.*;

/**
 * Created by yjl on 2018-07-30.
 */
public class GuiExcel extends JFrame {
    private String uploadPath = "E:\\ouDataTemp"; // 上传文件的目录
    //String imgPath=getRealPath().replace("\\","/")+"/resources"; //图片路径，开发环境用这个路径
    String imgPath = "E:/ouDataTemp";
    /**如果没打包后运行则debug为true*/
    public static boolean debug = false;

    JLabel l1 = new JLabel("please choose a data file..."); //提示信息
    JLabel pathLabel = new JLabel(""); //显示选中的文件路径
    //JTextArea westArea = new JTextArea(10, 10);  //TextArea：文本域，可以写好多行，10,10表示有10行10列
    //JTextArea centerArea = new JTextArea(10, 10);
    //JTextArea eastArea = new JTextArea(10, 10);
    JLabel logoImg = new JLabel(new ImageIcon(imgPath+"/images/logo.jpg"));
    JLabel nameImg = new JLabel(new ImageIcon(imgPath+"/images/name.jpg"));
    JPanel imgPanel = new JPanel(); //放置图片panel
    JPanel panel = new JPanel();    //放置选择文件按钮panel
    JPanel btnPanel = new JPanel();  //放置保存和重置按钮panel

    //JButton btn1 = new JButton("west--->center");
    JButton btn2 = new JButton("choose file");
    JButton saveBtn = new JButton("save");
    JButton resetBtn = new JButton("reset");

    //JFileChooser btn2 = new JFileChooser();
    String filePath=""; //模板文件目录
    String dirpath;    //文件所在路径
    String fileName;    //文件名称
    private FileDialog openDia, saveDia;// 定义“打开、保存”对话框

   // JButton btn3 = new JButton("center<----east");
    class Listener implements ActionListener {
       @Override
       public void actionPerformed(ActionEvent e) {
           Object source = e.getSource();
           if (source == btn2) {
               openDia.setVisible(true);//显示打开文件对话框

               dirpath = openDia.getDirectory();//获取打开文件路径并保存到字符串中。
               fileName = openDia.getFile();//获取打开文件名称并保存到字符串中
               pathLabel.setText(dirpath+fileName);
           }
           if (source==saveBtn){
               if ("".equals(pathLabel.getText())||null==pathLabel.getText()){
                   JOptionPane.showMessageDialog(null, "please choose a file", "warning\n", JOptionPane.ERROR_MESSAGE);
                   return;
               }
               cloneXls(fileName,dirpath);
               createXls();

               JOptionPane.showInternalMessageDialog(null, "information","information", JOptionPane.INFORMATION_MESSAGE);
           }
           if (source==resetBtn){
               pathLabel.setText("");
           }

       }
   }
    public GuiExcel(String title) {
        super(title);
        init();
    }
    public void init(){
        this.setLayout(new BorderLayout(10, 0));
        this.setTitle("Excel Generator by Mr. ou");
        l1.setFont(new Font("宋体",Font.BOLD,10));
        l1.setBackground(Color.GREEN);
        l1.setOpaque(true);//设置背景不透明
        //this.add(l1,BorderLayout.WEST);
        //this.add(westArea,BorderLayout.WEST);
        //this.add(centerArea);//没有添加的，默认是在中间位置
        //this.add(eastArea,BorderLayout.EAST);
        openDia = new FileDialog(this, "打开", FileDialog.LOAD);
        saveDia = new FileDialog(this, "保存", FileDialog.SAVE);

        // 嵌套布局
        imgPanel.setLayout(new FlowLayout(
                FlowLayout.CENTER,10,10)); //构建了FlowLayout的布局，居中，宽60，高20
        panel.setLayout(new GridLayout(3, 2,20,10));
        btnPanel.setLayout(new FlowLayout(
                FlowLayout.CENTER,10,100));
        //panel.add(btn1);
        imgPanel.add(logoImg);
        imgPanel.add(nameImg);
        panel.add(l1);
        panel.add(btn2);
        panel.add(pathLabel);
        btnPanel.add(saveBtn);
        btnPanel.add(resetBtn);
        //panel.add(btn3);
        this.add(imgPanel,BorderLayout.NORTH);
        this.add(panel,BorderLayout.CENTER);
        this.add(btnPanel,BorderLayout.SOUTH);
        ActionListener ac = new Listener();
        //btn1.addActionListener(ac);
        btn2.addActionListener(ac);
        saveBtn.addActionListener(ac);
        resetBtn.addActionListener(ac);
        //btn3.addActionListener(ac);

        this.setBounds(100, 100, 600, 500);
        this.setVisible(true);
        this.setDefaultCloseOperation(EXIT_ON_CLOSE);
    }

    //克隆一份选中的数据文件到E盘的uploadPath里，
    public void cloneXls(String fileName,String dirpath){
        File file = new File(uploadPath+"/allexcel");
        //如果目录不存在
        if (!file.isDirectory()) {
            //创建文件上传目录
            file.mkdirs();
        }
        File dataFile = new File(dirpath+fileName);
        InputStream in = null;
        OutputStream out = null;
        try {
             in = new FileInputStream(dataFile);
             out = new FileOutputStream(uploadPath+"/"+fileName); //名字在这里
            BufferedOutputStream bout = new BufferedOutputStream(out);
            int tempbyte;
            while ((tempbyte = in.read()) != -1) {
                System.out.write(tempbyte);
                bout.write(tempbyte);
            }
            //刷新缓冲区
            bout.flush();
        }catch (Exception e){
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

    public void createXls(){
        try{
            //filePath=getRealPath().replace("\\","/")+"/resources"; //开发环境用这个路径
            filePath = "E:/ouDataTemp"; //模板路径
        }catch(Exception e){
            e.printStackTrace();
        }


        Map<String,Object> map=new HashMap<String,Object>();
        //这个是模板文件
        //String muBanPath = "E:\\000.xls";
        Physical physical=new Physical();

        // 这个是excel数据文件：eg：E:/yujianlin/xxxxx.xls，必须是xls格式
        //如果后缀不是xls，请在excel另存为选择 xls 即可
        File file = new File(uploadPath+"/"+fileName);
        if(file==null||"".equals(file)){
            System.out.println("The file does not exist!please check again.");
            return;
        }
        //obj.readExcel(file);
        InputStream in = null;
        OutputStream out = null;
        try {
            //得到excel数据文件的所有数据
            java.util.List<java.util.List<String>> allData = GetExcelInfo.readExcel2(file);
            System.out.println(allData);

            for (int i =0;i<allData.size();i++){
                //得到每一条数据
                java.util.List<String> eachData = allData.get(i);

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
                //CPR-SEASON（第一列）-ESQ-VENDOR（第8）-BOARD / MODEL（第2）
                out = new FileOutputStream(uploadPath+"/allexcel/"+"CPR-"+eachData.get(0)+"-ESQ-"+eachData.get(7)+"-"+eachData.get(1)+".xls"); //名字在这里
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

    public static String getRealPath() {
        String realPath = GuiExcel.class.getClassLoader().getResource("").getFile();
        java.io.File file = new java.io.File(realPath);
        realPath = file.getAbsolutePath();//去掉了最前面的斜杠/
        try {
            realPath = java.net.URLDecoder.decode(realPath, "utf-8");
        } catch (Exception e) {
            e.printStackTrace();
        }
        return realPath;
    }

    private static String initProjectPathAndDebug(){
        java.net.URL url = GuiExcel.class.getProtectionDomain().getCodeSource().getLocation();
        String filePath = null;
        try {
            filePath = java.net.URLDecoder.decode(url.getPath(), "utf-8");
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        if (filePath.endsWith(".jar"))  {
            filePath = filePath.substring(0, filePath.lastIndexOf("/") + 1);
        }
        //如果以bin目录接，则说明是开发过程debug测试查询，返回上一层目录
        if (filePath.endsWith("bin/") || filePath.endsWith("bin\\") )  {
            debug = true;
            filePath = filePath.substring(0, filePath.lastIndexOf("bin"));
        }
        java.io.File file = new java.io.File(filePath);
        filePath = file.getAbsolutePath();
        return filePath;
    }


    public static void main(String[] args) {
        new GuiExcel("BorderLayOut...");
    }

}
