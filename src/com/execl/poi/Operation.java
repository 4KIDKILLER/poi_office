package com.execl.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 操作Excel 操作excel比较繁琐，正在努力学习中
 * 
 * @author Xeelong
 *
 */
public class Operation {
	
	public static void main(String[] args) throws Exception {
		Operation obj = new Operation();
		obj.opreation("E:\\WorkFile\\Java\\Test\\execl.xls");
	}
	
	public void opreation(String path) throws Exception {
		/**
		 * 操作.xls文件使用HSSF包
		 */
		//创建一个工作簿
		Workbook doc = new HSSFWorkbook(); 
		File file = new File(path);
		if(file.exists()) {
			//若存在重复文件则删除
			file.delete();
		}
		//创建第一个Sheet页
		doc.createSheet("One");
		//创建第二个Sheet页
		Sheet sheetTow = doc.createSheet("Two");
		//在第二个sheet页上创建一个单元格
		Row rowTow = sheetTow.createRow(0);//创建第一行
		rowTow.createCell(0).setCellValue(1);//在第一行中创建第一个单元格，值为int类型
		rowTow.createCell(1).setCellValue("String");//在第一行中创建第二个单元格，值为String类型
		rowTow.createCell(2).setCellValue(true);//在第第一行中创建第3个单元格，值为boolean类型
		//输出到指定位置
		FileOutputStream output = new FileOutputStream(file); 
		//输出
		doc.write(output);
		output.close();
	}
}