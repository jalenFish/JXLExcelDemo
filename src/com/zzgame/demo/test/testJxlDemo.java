package com.zzgame.demo.test;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.zzgame.demo.beans.Physical;
import com.zzgame.demo.util.ExcelTool;
import com.zzgame.demo.util.GetExcelInfo;

public class testJxlDemo {
	public static void main(String[] args) throws IOException {
		 Map<String,Object> map=new HashMap<String,Object>();
		 	//�����ģ���ļ�
	        String path = "E:\\000.xls";
			Physical physical=new Physical();

	        /*physical.setCompany("xxxxx���޹�˾");
	        physical.setDateTimes("2013-10-13");
//	        String area= hnabaseCityBO.findoneById(jsonObj.get("area").toString()).getBaseChn();
	        physical.setArea("Сկ");//
	        physical.setPlanNumber("60");
	        physical.setRealNumber("50");
	        physical.setPassNumber("40");
	        physical.setFailNumber("10");
	        physical.setEndYield("80%");
	        physical.setAssistOrg("21");
	        physical.setAssistPrice("21");
	        physical.setPrevCompany("xxxxx������");
	        physical.setPrevPrice("2");
	        physical.setDoctors("��ҽ��");//
	        physical.setComment("ע����Ϣxxxxx");*/
//	        physical.setMill("yujianlin");
//	        physical.setModel("lalalla");
//	        physical.setSeason("spring");
//	        map.put("physical", physical);

			// �����excel�����ļ���eg��E:/yujianlin/xxxxx.xls��������xls��ʽ
			//�����׺����xls������excel���Ϊѡ�� xls ����
			File file = new File("E:/111.xls");
			if(file==null||"".equals(file)){
				System.out.println("�����ļ������ڣ������ļ��Ƿ����");
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
					map.put("physical", physical);

					in = ExcelTool.exportExcel(path, map);
					 out = new FileOutputStream("E:/ouData/�ཨ��"+i+".xls");
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
}
