package com.mrllking;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainApp {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		String file = "Bao_Cao_DK01TB_94TTT_944HH_31603_20160131.xlsx";
		String input = "E:\\input\\"+file;
		String output = "E:\\output\\"+file;
		
		System.out.println("==START===");
		
		Workbook inputWB = new XSSFWorkbook(new FileInputStream(input));
		System.out.println("==...===");
		Workbook outputWB = new XSSFWorkbook();
		
		Sheet inputSheet = inputWB.getSheet("DK01TB");
		Sheet outputSheet = outputWB.createSheet("ABC");
		
		Iterator<Row> inputRow = inputSheet.iterator();
		String maHo;
		String soHoKhau;
		String diaChi;
		Row rowTemp;
		int index = 1;
		System.out.println("Start reading...");
		while (inputRow.hasNext()) {
			Row row = inputRow.next();
			if(getTextValue(row.getCell(1)).equals("1")) {
				// begin a family
				// Get ma ho
				rowTemp = inputSheet.getRow(row.getRowNum()-6);
				maHo = getTextValue(rowTemp.getCell(8));
				maHo = maHo.substring(7);
				// Get soHoKhau
				rowTemp = inputSheet.getRow(row.getRowNum()-5);
				soHoKhau = getTextValue(rowTemp.getCell(1));
				soHoKhau = soHoKhau.substring(33);
				// Get dia Chi
				rowTemp = inputSheet.getRow(row.getRowNum()-3);
				diaChi = getTextValue(rowTemp.getCell(1));
				diaChi = diaChi.substring(9);
				
				while(inputRow.hasNext()) {
					row = inputRow.next();
					if(getTextValue(row.getCell(1)).isEmpty()) {
						break;
					}
					
					insertRow(outputSheet, row, maHo, soHoKhau, diaChi, index);
					index++;
				}
			}
		}

		FileOutputStream fileOut = new FileOutputStream(output);
		outputWB.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}
	
	private static void insertRow(Sheet outputSheet, Row inputRow, String maHo, String soHoKhau, String diaChi, int index) {
		Row row = outputSheet.createRow(index);

		row.createCell(0).setCellValue(index);
		row.createCell(1).setCellValue(getTextValue(inputRow.getCell(2)));
		row.createCell(2).setCellValue(getTextValue(inputRow.getCell(4)));
		row.createCell(3).setCellValue(getTextValue(inputRow.getCell(5)));
		row.createCell(4).setCellValue(getTextValue(inputRow.getCell(6)));
		row.createCell(5).setCellValue(getTextValue(inputRow.getCell(7)));
		row.createCell(6).setCellValue(getTextValue(inputRow.getCell(9)));
		row.createCell(7).setCellValue(getTextValue(inputRow.getCell(10)));
		row.createCell(8).setCellValue(getTextValue(inputRow.getCell(13)));
		row.createCell(9).setCellValue(maHo);
		row.createCell(10).setCellValue(soHoKhau);
		row.createCell(11).setCellValue(diaChi);
	}


	private static String getTextValue(Cell cell) {
        if (cell == null)
            return "";
        //cell.setCellType(Cell.CELL_TYPE_STRING);
        return cell.getStringCellValue();
    }
}
