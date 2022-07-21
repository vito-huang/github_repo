update﻿ infromation
package com.management;

import java.io.File;
import java.io.FileInputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.poi.readExcel.ReadExcelUtils;
import com.poi.readExcel.entity.DataEntity;
import com.poi.readExcel.entity.DataSheetEntity;
import com.poi.readExcel.entity.ExcelEntity;
import com.poi.readExcel.entity.TableEntity;
import com.poi.readExcel.entity.TableSheetEntity;

public class Start {

	public static void main(String[] args) {
		
		
		//	每5秒執行一次
		int intervalMs = 5000;
		// thread最大數量
		int threadMaxCount = 5;
		
		Executor executor = Executors.newFixedThreadPool(threadMaxCount);  
		executor.execute(new Work(intervalMs));  
		
		
	}


	public ExcelEntity readExcel(String path) {
		
		ExcelEntity ee = new ExcelEntity();
		String sheetName = "";
		
		File file = new File(path);
		List<String> errorMessage = new ArrayList<String>();
		try (
				FileInputStream fs = new FileInputStream(file);
				//.xlsx
				XSSFWorkbook workbook = new XSSFWorkbook(fs);
			) {
				Integer sheetCount = workbook.getNumberOfSheets();
				
				List<TableSheetEntity> tseList = new ArrayList<TableSheetEntity>();
				List<DataSheetEntity> dseList = new ArrayList<DataSheetEntity>();
				for(int i = 0 ; i < sheetCount ; i+=1) {
					XSSFSheet sheet = workbook.getSheetAt(i);
					sheetName = sheet.getSheetName();
					String[] sheetNameArray = sheetName.split("_");
					String sheetType = sheetNameArray[0];
					String tableName = sheetNameArray[1];
					if(sheetType.equals("table")) {
						TableSheetEntity tse = readTableSheet(sheet, tableName);
						tseList.add(tse);
					} else if (sheetType.equals("data")) {
						DataSheetEntity dse = readDataSheet(sheet, tableName, tseList);
						dseList.add(dse);
					}
				}
				System.out.println("tseList:"+tseList.size());
				System.out.println("dseList:"+dseList.size());
				ee.setTableSheetEntityList(tseList);
				ee.setDataSheetEntityList(dseList);
			} catch (Exception e) {
				e.printStackTrace();
				errorMessage.add(e.getMessage());
			}
		return ee;
	}
	

	
	
	public TableSheetEntity readTableSheet(XSSFSheet sheet, String tableName) {
		ReadExcelUtils utils = new ReadExcelUtils();
		TableSheetEntity tse = new TableSheetEntity();
		tse.setTableName(tableName);
		Iterator<Row> rowIterator = sheet.iterator();
		List<TableEntity> tableEntityList = new ArrayList<TableEntity>();
		
		int i = 0;
		while (rowIterator.hasNext()){
			Row row = rowIterator.next();
			if (i != 0) {
				TableEntity te = utils.getTableEntity(row);
				tableEntityList.add(te);
			}
			i += 1;
		}
		
		tse.setTableEntityList(tableEntityList);
		return tse;
	}
	
	public DataSheetEntity readDataSheet(XSSFSheet sheet, String tableName, List<TableSheetEntity> tseList) {
		ReadExcelUtils utils = new ReadExcelUtils();
		DataSheetEntity dse = new DataSheetEntity();
		dse.setTableName(tableName);
		Iterator<Row> rowIterator = sheet.iterator();
		List<DataEntity> dataEntityList = new ArrayList<DataEntity>();
		
		
		TableSheetEntity sameTse = tseList.stream()
				.filter(tse -> tse.getTableName().equals(tableName)).findFirst().orElse(null);
		
		List<TableEntity> headerColumnOrderList = new ArrayList<TableEntity>();
		
		int i = 0;
		while (rowIterator.hasNext()){
			Row row = rowIterator.next();
			if (i == 0) {
				headerColumnOrderList = utils.accordingToDataHeaderOrder(row, sameTse);
			} else {
				DataEntity de = utils.getDataEntity(row, headerColumnOrderList);
				dataEntityList.add(de);
			}
			i += 1;
		}
		dse.setDataEntityList(dataEntityList);
		return dse;
	}
	
	public TableEntity getTableEntity(Row row) {
		TableEntity te = new TableEntity();
		
		for (int i = 0; i < columnNamrArray.length ; i += 1) {
			Cell cell = row.getCell(i);
			if (cell == null) {
				continue;
			}
			String content = cell.getStringCellValue().trim();
			
			if ( i == 0) {
				//	column
				te.setColum(content);
			} else if (i == 1) {
				// type
				te.setType(content);
			} else if (i == 2) {
				// pk

				Boolean pk = content.equals("y")? true: false;
				te.setPk(pk);
			} else if (i == 3) {
				// null
				Boolean notNull = content.equals("y")? true: false;
				te.setNotNull(notNull);
			}
		}
		return te;
	}
	
	/**
	 * 
	 * @param dataRow
	 * @param headerColumnOrderList �ھ�data sheet��header���ǭ��s�ƹL��
	 * @return
	 */
	public DataEntity getDataEntity(Row dataRow, List<TableEntity> headerColumnOrderList) {
		Iterator<Cell> dataIterator = dataRow.iterator();
		
		DecimalFormat intFormat = new DecimalFormat("0.#");
		
		
		DataEntity de = new DataEntity();
		Map<String, String> map = new HashedMap<String, String>();
		int i = 0;
		while(dataIterator.hasNext()) {
			
			Cell dataCell = dataIterator.next();
			TableEntity te = headerColumnOrderList.get(i);
			String columnType = te.getType();
			String columnName = te.getColum();
			
			String value = "";
			if (columnType.startsWith("int")) {
				value = String.valueOf(intFormat.format(dataCell.getNumericCellValue()));
			} else if (columnType.startsWith("nvarchar")) {
				value = dataCell.getStringCellValue();
			}
            map.put(columnName, value);
			i += 1;
		}
		de.setMap(map);
		return de;
	}
	
	public List<TableEntity> accordingToDataHeaderOrder(Row dataHeaderRow, TableSheetEntity sameTse) {
		List<TableEntity> headerColumnOrderList = new ArrayList<TableEntity>();
		
		Iterator<Cell> headerCellIterator = dataHeaderRow.cellIterator();
		//	�ھ�data sheet��header column���ǨӮھ�Ū����Ʈɪ����
		while(headerCellIterator.hasNext()) {
			Cell cell = headerCellIterator.next();
			String columnName = cell.getStringCellValue().trim();
			
			TableEntity sameTe = sameTse.getTableEntityList().stream()
				.filter(te -> te.getColum().equalsIgnoreCase(columnName)).findFirst().orElse(null);
			headerColumnOrderList.add(sameTe);
		}
		return headerColumnOrderList;
	}

	
	
	public static String getErrorMessage(Exception e, String sheetName) {
		String str = String.format("sheetName:%s,message:%s", sheetName, e.getMessage());
		return str;
	}
		
	
	
	
}
