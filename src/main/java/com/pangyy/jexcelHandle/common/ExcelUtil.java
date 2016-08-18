package com.pangyy.jexcelHandle.common;

/**
 * Created by pangyaoyang on 2016/6/15.
 * 读写Excel 文件，支持office2003的xls文件和 office2007的xlsx文件。
 */
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

public class ExcelUtil {

	public static boolean isXlsFile(String fileName) {
		String excel2003Suffix = ".xls";
		return (fileName.indexOf(excel2003Suffix) == fileName.length()
				- excel2003Suffix.length());
	}

	public static boolean isXlsxFile(String fileName) {
		String excel2007Suffix = ".xlsx";
		return (fileName.indexOf(excel2007Suffix) == fileName.length()
				- excel2007Suffix.length());
	}

	/**
	 * 检查是否是excel文件
	 * 
	 * @param fileName  文件名
	 * @return true/false
	 */
	public static boolean isExcelFile(String fileName) {

		if (StringUtils.isBlank(fileName)) {
			return false;
		}
		return isXlsFile(fileName) || isXlsxFile(fileName);
	}

	/**
	 * 判断是否为空行
	 * 
	 * @param row 行对象
	 * @return true 空 false 非空
	 */
	public static boolean isEmptyRow(Row row) {
		if (row == null) {
			return true;
		}

		boolean result = true;

		for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {

			Cell cell = row.getCell(i, HSSFRow.RETURN_BLANK_AS_NULL);
			String value = "";
			if (cell != null) {
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					value = cell.getStringCellValue();
					break;
				case Cell.CELL_TYPE_NUMERIC:
					value = String.valueOf((int) cell.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					value = String.valueOf(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_FORMULA:
					value = String.valueOf(cell.getCellFormula());
					break;
				default:
					break;
				}

				if (StringUtils.isNotBlank(value.trim())) {
					result = false;
					break;
				}
			}
		}

		return result;
	}

	/**
	 * 读取excel文件内容到数组。
	 * 
	 * @param fileName   excel文件名
	 * @return 每行数据放入一个数组A，多行数据生成的多个数组A再放入一个数组B，即数组的数组 B[A[]]
	 */
	public static ArrayList<Object> readXlsxFileToArray(String fileName) {

		if (!isExcelFile(fileName)) {
			System.out.println("readXlsxFileToArray: " + "不是excel文件");
			return null;
		}

		ArrayList<Object> result = new ArrayList<Object>();

		InputStream stream = null;
		try {
			stream = new FileInputStream(fileName);
			Workbook wb = null;
			if (isXlsFile(fileName)) {
				wb = new HSSFWorkbook(stream);
			} else if (isXlsxFile(fileName)) {
				wb = new XSSFWorkbook(stream);
			}
			if (wb == null) {
				System.out.println("readXlsxFileToArray: " + "文件打开失败");
				return null;
			}
			Sheet sheet1 = wb.getSheetAt(0);

			int maxCellNum = 0;
			for (int i = 0; i <= sheet1.getLastRowNum(); i++) {
				Row row = sheet1.getRow(i);
				if (row == null || isEmptyRow(row)) {
					break;
				}
				
				/**
				 * 最大列数由第一行列数决定,因为一般第一行为标题，后续行的列里面有空列
				 */
				if (i == 0) {
					maxCellNum = row.getLastCellNum();
				}
				ArrayList<String> cellResult = new ArrayList<String>();

				for (int j = 0; j < maxCellNum; j++) {
					Cell cell = row.getCell(j);
					String value = "";

					if (cell == null) {
						cellResult.add(value);
						continue;
					}
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						value = cell.getStringCellValue();
						break;
					case Cell.CELL_TYPE_NUMERIC:
						value = String
								.valueOf((int) cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						value = String.valueOf(cell.getBooleanCellValue());
						break;
					case Cell.CELL_TYPE_FORMULA:
						value = String.valueOf(cell.getCellFormula());
						break;
					case Cell.CELL_TYPE_BLANK:
						value = "";
						break;
					case Cell.CELL_TYPE_ERROR:
						System.out.println("readXlsxFileToArray: " + "错误单元格");
						return null;
					default:
						break;
					}
					cellResult.add(value);
				}
				result.add(cellResult);
			}
		} catch (Exception e) {
			System.out.println("readXlsxFileToArray: " + e.getMessage());
		} finally {
			try {
				if (stream != null) {
					stream.close();
				}
			} catch (IOException e) {
				System.out.println("[readXlsxFileToArray]：关闭excel文件流异常:"
						+ e.getMessage());
			}
		}

		System.out.println("[readXlsxFileToArray] 完成");
		return result;
	}

}
