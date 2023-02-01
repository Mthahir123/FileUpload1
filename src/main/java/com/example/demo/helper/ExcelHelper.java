package com.example.demo.helper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.QuoteMode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.entity.Tutorial;

public class ExcelHelper {

	public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	static String[] HEADERs = { "Id", "Title", "Description", "Published" };
	static String SHEET = "Tutorials";

	public static boolean hasExcelFormat(MultipartFile file) {

		if (!TYPE.equals(file.getContentType())) {
			return false;
		}

		return true;
	}

	public static List<Tutorial> excelToTutorials(InputStream is) {
		try {
			Workbook workbook = new XSSFWorkbook(is);
			Sheet sheet = workbook.getSheet(SHEET);
			Iterator<Row> rows = sheet.iterator();

			List<Tutorial> tutorials = new ArrayList<Tutorial>();

			int rowNumber = 0;
			while (rows.hasNext()) {
				Row currentRow = rows.next();
				if (rowNumber == 0) {
					rowNumber++;
					continue;
				}
				Iterator<Cell> cellsInRow = currentRow.iterator();
				Tutorial tutorial = new Tutorial();
				int cellIdX = 0;
				while (cellsInRow.hasNext()) {
					Cell currentCell = cellsInRow.next();

					switch (cellIdX) {
					case 0:
						tutorial.setId((long) currentCell.getNumericCellValue());
						break;
					case 1:
						tutorial.setTitle(currentCell.getStringCellValue());
						break;

					case 2:
						tutorial.setDescription(currentCell.getStringCellValue());
						break;

					case 3:
						tutorial.setPublished(currentCell.getBooleanCellValue());
						break;

					default:
						break;
					}
					cellIdX++;
				}
				tutorials.add(tutorial);
			}
			workbook.close();
			return tutorials;
		} catch (IOException e) {
			throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
		}
	}
	
	public static ByteArrayInputStream tutorialsToExcel(List<Tutorial> tutorials) {
		try {
			Workbook workbook = new XSSFWorkbook();
			Sheet sheet = workbook.createSheet(SHEET);
			int n=0;
			Row headRow = sheet.createRow(n++);
			for(int i=0;i<HEADERs.length;i++) {
				Cell cell = headRow.createCell(i);
				cell.setCellValue(HEADERs[i]);
			}
			for(Tutorial tutorial : tutorials) {
				Row bodyRow = sheet.createRow(n++);
				Cell cell = bodyRow.createCell(0);
				cell.setCellValue(tutorial.getId());
				cell = bodyRow.createCell(1);
				cell.setCellValue(tutorial.getTitle());
				cell = bodyRow.createCell(2);
				cell.setCellValue(tutorial.getDescription());
				cell = bodyRow.createCell(3);
				cell.setCellValue(tutorial.isPublished());
			}
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			workbook.write(out);
			workbook.close();
			out.close();
			return new ByteArrayInputStream(out.toByteArray());
		} catch (IOException e) {
			throw new RuntimeException("fail to import data to CSV file: " + e.getMessage());
		}
	}
}
