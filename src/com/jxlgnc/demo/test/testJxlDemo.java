package com.jxlgnc.demo.test;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.jxlgnc.demo.beans.Physical;
import com.jxlgnc.demo.util.ExcelTool;
import com.jxlgnc.demo.util.GetExcelInfo;

public class testJxlDemo {
	public static void main(String[] args) throws IOException {
		 Map<String,Object> map=new HashMap<String,Object>();
		 	//这个是模板文件
	        String path = "E:\\000.xls";
			Physical physical=new Physical();

			// 这个是excel数据文件：eg：E:/yujianlin/xxxxx.xls，必须是xls格式
			//如果后缀不是xls，请在excel另存为选择 xls 即可
			File file = new File("E:/111.xls");
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

					in = ExcelTool.exportExcel(path, map);
					 out = new FileOutputStream("E:/ouData/余建林"+i+".xls");
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
}
