package cn.jerry.mouse.property_tools;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExcelProperties {

	public void exportToFile(String excelFileName,int keyColumn,int valueColumn,String exportFileName) throws Exception {
		File excelFile = new File(excelFileName);
		File exportFile = new File(exportFileName);
		BufferedWriter bw = new BufferedWriter(new FileWriter(exportFile));
		
		Workbook workbook = Workbook.getWorkbook(excelFile);
		Sheet[] sheet = workbook.getSheets();
		Cell[] keyCell;
		Cell[] valueCell;
		String key;
		String value;
		
		for(int sheetIndex=0;sheetIndex<sheet.length;sheetIndex++)
		{
			keyCell = sheet[sheetIndex].getColumn(keyColumn);
			valueCell = sheet[sheetIndex].getColumn(valueColumn);
			for(int rowIndex=1;rowIndex<keyCell.length;rowIndex++) //第一行作为标题栏，忽略掉
			{
				key = keyCell[rowIndex].getContents();
				value = valueCell[rowIndex].getContents();
				if(key.trim()!="")
					if(valueColumn == 1){
						value = java.net.URLEncoder.encode(value,"utf-8");
					}
					bw.write(key+"="+value+"\r\n");
			}
		}
		bw.close();
		workbook.close();
	}
	private boolean isFileValid(String excelFileName) throws Exception
	{
		try {
			File excelFile = new File(excelFileName);
			Workbook workbook = Workbook.getWorkbook(excelFile);
			workbook.getSheets();
			workbook.close();
		} catch (BiffException e) {
			throw new Exception("不支持此文件格式，仅支持Excel 2003");
		}
		return true; 
	}
	public static void main(String[] args) throws Exception
	{
		ExcelProperties excelUtil = new ExcelProperties();
		String excelFilePath = "D:\\res.xls";
		String exportCNFilePath = "D:\\res_zh_CN.properties";
		String exportENFilePath = "D:\\res_en_US.properties";
		String exportDefaultFilePath = "D:\\res.properties";
		int keyColumn = 0;
		int valueColumn;
		
		if(args.length==4)
		{
			excelFilePath = args[0];
			exportCNFilePath = args[1];
			exportENFilePath = args[2];
			exportDefaultFilePath = args[3];
		}
		else if(args.length!=0)
		{
			System.out.println("Usage: java -jar ExcelProps.jar excelFilePath exportCNFilePath exportENFilePath excelFilePath");
			return;
		}
		
		excelUtil.isFileValid(excelFilePath);
		
		System.out.println("Begin to exprort from excelFile: "+excelFilePath);
		
		valueColumn = 1;
		excelUtil.exportToFile(excelFilePath, keyColumn, valueColumn, exportCNFilePath);
		System.out.println("Config file in Chinese exported: "+exportCNFilePath);
		
		valueColumn = 2;
		excelUtil.exportToFile(excelFilePath, keyColumn, valueColumn, exportENFilePath);
		System.out.println("Config file in English exported: "+exportENFilePath);
		
		valueColumn = 1;
		excelUtil.exportToFile(excelFilePath, keyColumn, valueColumn, exportDefaultFilePath);
		System.out.println("Config file in default language exported: "+exportDefaultFilePath);
	}
}
