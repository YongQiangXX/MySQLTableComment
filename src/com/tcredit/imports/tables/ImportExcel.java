package com.tcredit.imports.tables;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ImportExcel {

	public static void importExcel(Map<String, List<String>> map) throws Exception {
		/**
		 * 注意这只是07版本以前的做法对应的excel文件的后缀名为.xls 07版本和07版本以后的做法excel文件的后缀名为.xlsx
		 */
		// 创建新工作簿
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 新建工作表
		HSSFSheet sheet = workbook.createSheet("risk_data_two_sys");
		// 创建行,行号作为参数传递给createRow()方法,第一行从0开始计算
		Set<Entry<String, List<String>>> entrySet = map.entrySet();
		int j = 1;
		for (Entry<String, List<String>> entry : entrySet) {
			List<String> list = entry.getValue();
			String key = entry.getKey();
			for (int i = 0; i < list.size(); i++) {
				String table = list.get(i);
				String[] ta = table.split("--");
				HSSFRow row = sheet.createRow(j);
				if (ta[0].equals("id")) {
					HSSFCell cell1 = row.createCell(0);
					cell1.setCellValue(key);
				}

				HSSFCell cell1 = row.createCell(1);
				cell1.setCellValue(ta[0]);
				HSSFCell cell2 = row.createCell(2);
				cell2.setCellValue(ta[1]);

				if (ta.length == 3) {
					HSSFCell cell4 = row.createCell(3);
					cell4.setCellValue(ta[2]);
				}
				j++;
			}
		}
		FileOutputStream fos = new FileOutputStream(new File("/java/1.xlsx"));
		workbook.write(fos);
		fos.close();
		System.out.println("*****数据导入成功*****");
	}
}
