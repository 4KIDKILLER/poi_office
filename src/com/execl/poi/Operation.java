package com.execl.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * ����Excel ����excel�ȽϷ���������Ŭ��ѧϰ��
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
		 * ����.xls�ļ�ʹ��HSSF��
		 */
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
		rowTow.createCell(2).setCellValue(true);//�ڵڵ�һ���д�����3����Ԫ��ֵΪboolean����
		//�����ָ��λ��
		FileOutputStream output = new FileOutputStream(file); 
		//���
		doc.write(output);
		output.close();
	}
}