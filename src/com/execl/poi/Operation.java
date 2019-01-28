package com.execl.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 操作Excel 操作excel比较繁琐，正在努力学习中
 * 
 * @author Xeelong
 *
 */
public class Operation {
	
	public static void main(String[] args) throws Exception {
		Operation obj = new Operation();
//		obj.createXls("E:\\WorkFile\\Java\\Test\\execl.xls");
		obj.createXlsx("E:\\WorkFile\\Java\\Test\\execl.xlsx");
	}
	/**
	 * 创建.xls文件使用HSSF包
	 */
	public void createXls(String path) throws Exception {
		String suffix = path.substring(path.lastIndexOf("."));
		if(!suffix.equals(".xls")) {
			System.out.println("请输入.xls文件");
			return;
		}
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
		rowTow.createCell(2).setCellValue(true);//在第一行中创建第3个单元格，值为boolean类型
		//输出到指定位置
		FileOutputStream output = new FileOutputStream(file); 
		//输出
		doc.write(output);
		output.close();
	}

	/**
	 * 创建.xlsx文件使用XSSF包
	 */
	public void createXlsx(String path) throws Exception {
		String suffix = path.substring(path.lastIndexOf("."));
		if(!suffix.equals(".xlsx")) {
			System.out.println("请输入.xlsx文件");
			return;
		}
		Workbook doc = new XSSFWorkbook();
		File file = new File(path);
		if(file.exists()) {
			//若存在重复文件则删除
			file.delete();
		}
		//创建工作簿
		Sheet sheetOne = doc.createSheet("One");
		//对第一个sheet进行操作
		Row row1 = sheetOne.createRow(0);//创建行1
			//往第一行添加单元格并设置值
			row1.createCell(0).setCellValue(1);
			row1.createCell(1).setCellValue("String");
			row1.createCell(2).setCellValue(false);
		Row row2 = sheetOne.createRow(1);//创建行2
			//往第一行添加单元格并设置值
			row2.createCell(0).setCellValue(false);
			row2.createCell(1).setCellValue("String");
			row2.createCell(2).setCellValue(1);
		doc.createSheet("Two");
		FileOutputStream output = new FileOutputStream(file);
		doc.write(output);
		output.close();
	}

}