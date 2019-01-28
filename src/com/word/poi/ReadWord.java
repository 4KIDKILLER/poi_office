package com.word.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * ��ȡword�ĵ�
 * 
 * for POI
 * 
 * @author Xeelong
 *
 */
public class ReadWord {
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		ReadWord word = new ReadWord();
		String data = word.readDocx("E:\\WorkFile\\Java\\Test\\word.docx");
		System.out.println(data);
	}

	/**
	 * ��ȡdoc��ʽ��word�ĵ�
	 */
	public String readDoc(String path) {
		String suffix = path.substring(path.lastIndexOf("."));
		if (!suffix.equals(".doc")) {
			return "�ļ���ʽ������ѡ��.doc��ʽ�ĵ�";
		}
		File file = new File(path);
		FileInputStream read = null;
		String docText = "";
		try {
			read = new FileInputStream(file);
			WordExtractor result =  new WordExtractor(read);
			docText = result.getText();
			result.close();
			read.close();
		} catch (Exception e) {
			System.out.println("��ȡ�ļ�����");
			e.printStackTrace();
		}
		return docText;
	}
	/**
	 * ��ȡdocx��ʽ��word�ĵ� 
	 */
	public String readDocx(String path) {
		String suffix = path.substring(path.lastIndexOf("."));
		if(!suffix.equals(".docx")) {
			return "�ļ���ʽ������ѡ��.docx��ʽ�ĵ�";
		}
		File file = new File(path);
		FileInputStream read = null;
		XWPFWordExtractor result = null;
		try {
            read = new FileInputStream(file);
            XWPFDocument doc = new XWPFDocument(read);  
            result = new XWPFWordExtractor(doc);  
		} catch (Exception e) {
			System.out.println("��ȡ�ļ�����");
			e.printStackTrace();
		}
		return result.getText();
	}
}
