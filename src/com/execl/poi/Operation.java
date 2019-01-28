package com.execl.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * ����Excel ����excel�ȽϷ���������Ŭ��ѧϰ��
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
	 * ����.xls�ļ�ʹ��HSSF��
	 */
	public void createXls(String path) throws Exception {
		String suffix = path.substring(path.lastIndexOf("."));
		if(!suffix.equals(".xls")) {
			System.out.println("������.xls�ļ�");
			return;
		}
		//����һ��������
		Workbook doc = new HSSFWorkbook(); 
		File file = new File(path);
		if(file.exists()) {
			//�������ظ��ļ���ɾ��
			file.delete();
		}
		//������һ��Sheetҳ
		doc.createSheet("One");
		//�����ڶ���Sheetҳ
		Sheet sheetTow = doc.createSheet("Two");
		//�ڵڶ���sheetҳ�ϴ���һ����Ԫ��
		Row rowTow = sheetTow.createRow(0);//������һ��
		rowTow.createCell(0).setCellValue(1);//�ڵ�һ���д�����һ����Ԫ��ֵΪint����
		rowTow.createCell(1).setCellValue("String");//�ڵ�һ���д����ڶ�����Ԫ��ֵΪString����
		rowTow.createCell(2).setCellValue(true);//�ڵ�һ���д�����3����Ԫ��ֵΪboolean����
		//�����ָ��λ��
		FileOutputStream output = new FileOutputStream(file); 
		//���
		doc.write(output);
		output.close();
	}

	/**
	 * ����.xlsx�ļ�ʹ��XSSF��
	 */
	public void createXlsx(String path) throws Exception {
		String suffix = path.substring(path.lastIndexOf("."));
		if(!suffix.equals(".xlsx")) {
			System.out.println("������.xlsx�ļ�");
			return;
		}
		Workbook doc = new XSSFWorkbook();
		File file = new File(path);
		if(file.exists()) {
			//�������ظ��ļ���ɾ��
			file.delete();
		}
		//����������
		Sheet sheetOne = doc.createSheet("One");
		//�Ե�һ��sheet���в���
		Row row1 = sheetOne.createRow(0);//������1
			//����һ����ӵ�Ԫ������ֵ
			row1.createCell(0).setCellValue(1);
			row1.createCell(1).setCellValue("String");
			row1.createCell(2).setCellValue(false);
		Row row2 = sheetOne.createRow(1);//������2
			//����һ����ӵ�Ԫ������ֵ
			row2.createCell(0).setCellValue(false);
			row2.createCell(1).setCellValue("String");
			row2.createCell(2).setCellValue(1);
		doc.createSheet("Two");
		FileOutputStream output = new FileOutputStream(file);
		doc.write(output);
		output.close();
	}

}