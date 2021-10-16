package com.Sayan.ExcelProject;

import java.io.File;
import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateInvoices {

	public static void main(String[] args) {
		
		try {
			
			//create .xlsx
			Workbook workbook = new XSSFWorkbook();
			
			//For Older .xls HSSFWorkbook
			
			//Create Sheet
			Sheet sh = workbook.createSheet("Invoices");
			
			//Header - Create Top Row With Column Headings
			String[] columnHeadings = {
					"Item id",
					"Item Name",
					"Quantity",
					"Item Price",
					"Sold Date"
			};
			//Heading Bold with Fore Color
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short)12);
			headerFont.setColor(IndexedColors.BLACK.index);
			
			//Create a cell style
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFont(headerFont);
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
			
			//Create Header Row
			Row headerRow = sh.createRow(0);
			//Loop over the Header Array and Create as you go
			for (int i = 0; i < columnHeadings.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columnHeadings[i]);
				cell.setCellStyle(headerStyle);
			}
			
			//Freeze Rows : Here Header Rows
			sh.createFreezePane(1, 1);
			
			//Header Done
			//==========================
			
			//Fill Data
			ArrayList<Invoices> a = createData();
			CreationHelper creationHelper = workbook.getCreationHelper();
			CellStyle dateStyle = workbook.createCellStyle();
			dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("MM/dd/yyyy")); 
			
			//Filling Rows
			int rownum = 1;
			for(Invoices i : a) {
				Row row = sh.createRow(rownum++);
				row.createCell(0).setCellValue(i.getItemId());
				row.createCell(1).setCellValue(i.getItemName());
				row.createCell(2).setCellValue(i.getItemQty());
				row.createCell(3).setCellValue(i.getTotalPrice());
				Cell dateCell = row.createCell(4);
				dateCell.setCellValue(i.getItemSoldDate());
				dateCell.setCellStyle(dateStyle);
			}
			
			//Auto-Size Columns
			for (int i = 0; i < columnHeadings.length; i++) {
				sh.autoSizeColumn(i);
			}
			
			//Create rows with Formula i.e. = SUM
			Row sumRow = sh.createRow(rownum);
			Cell sumRowTitle = sumRow.createCell(0);
			sumRowTitle.setCellValue("Total");
			sumRowTitle.setCellStyle(headerStyle);
			
			String strFormula = "SUM(D2:D"+rownum+")";
			Cell sumCell = sumRow.createCell(3);
			sumCell.setCellFormula(strFormula);
			sumCell.setCellValue(true);
			
			//Group Rows and Collapse Them
			int noOfRows = sh.getLastRowNum();
			sh.groupRow(1, noOfRows-1);
			sh.setRowGroupCollapsed(1, true);
			
			//New Sheet
			Sheet sh2 = workbook.createSheet("Second");
			
			//Writing The Output to a File
			String fileSeperator = File.separator;
			FileOutputStream fileOut = new FileOutputStream("dist"+fileSeperator+"Docs"+fileSeperator+"Invoice.xlsx");
			workbook.write(fileOut);			
			//Close The Streams
			fileOut.close();
			workbook.close();
			
			
			//Print - Confirmation of Job Done
			System.out.println("Job Done");
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}

	private static ArrayList<Invoices> createData() throws ParseException {
		ArrayList<Invoices> a = new ArrayList();
		
		//Adding Data
		a.add(new Invoices(1, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(2, "Table", 3, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
		a.add(new Invoices(3, "Hook", 4, 30.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/03/2020")));
		a.add(new Invoices(4, "Chair", 5, 40.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/04/2020")));
		a.add(new Invoices(5, "Cups", 6, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/05/2020")));
		a.add(new Invoices(6, "Plate", 7, 60.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/06/2020")));
		a.add(new Invoices(7, "Plate", 8, 90.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/06/2020")));
		
		return a;
	}

}
