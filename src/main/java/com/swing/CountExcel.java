package com.swing;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import javax.swing.JOptionPane;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CountExcel {

	private static int countExcel(String filePath, String folderPath, String fileName, int selectedCounted,
			String selectedCountedName, int sheetNum) {

//		filecreateFolder = new File(folderPath +);

		int count = 0;
		try {

			FileInputStream file1Count = new FileInputStream(filePath);
			XSSFWorkbook workBookCount = new XSSFWorkbook(file1Count);
			XSSFSheet sheetCount = workBookCount.getSheetAt(sheetNum);

			int totalNumberOfRowsInExcel1Count = sheetCount.getLastRowNum();

//			System.out.println("totalNumberOfRowsInExcel1Count:"+totalNumberOfRowsInExcel1Count);
			
			int columnIndex = selectedCounted;

			int total = 0;

			Set<String> set = new HashSet<>();

			for (int r = 1; r <= totalNumberOfRowsInExcel1Count; r++) {
//				System.out.println("r:"+r);
//				if (sheetCount.getRow(r) == null) {
//					continue;
//				} else if(sheetCount.getRow(r).getCell(columnIndex) == null) {
//					continue;
//				} 
				set.add(sheetCount.getRow(r).getCell(columnIndex).toString());
			}

			String[] setToStringArr = set.toArray(new String[set.size()]);
			int[] arr = new int[set.size()];

//			for (int i = 0; i < setToStringArr.length; i++) {
//				System.out.println(setToStringArr[i]);
//			}

			for (int i = 0; i < setToStringArr.length; i++) {

				for (int j = 1; j <= totalNumberOfRowsInExcel1Count; j++) {

					if (setToStringArr[i].equalsIgnoreCase(sheetCount.getRow(j).getCell(columnIndex).toString())) {
						arr[i]++;
					}
				}
			}

			for (int count1 : arr) {
//				System.out.println(count);
				total += count1;
			}

			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;

			for (int r = 0; r <= setToStringArr.length; r++) {
				rowCreated = sheetCreate1.createRow(r);

				for (int c = 0; c < 3; c++) {
					rowCreated.createCell(c);
				}
			}

			for (int c = 0; c < 3; c++) {
				for (int i = 0; i <= setToStringArr.length; i++) {

					if (i < setToStringArr.length) {

						if (c == 0 && i == 0) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(selectedCountedName);
						} else if (c == 1) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(setToStringArr[i]);
						} else if (c == 2) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(arr[i]);
						}

					}

					if (i <= setToStringArr.length) {

						if (c == 1 && i == setToStringArr.length) {
							sheetCreate1.getRow(i).getCell(c).setCellValue("total:");
						} else if (c == 2 && i == setToStringArr.length) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(total);
						}

					}
				}
			}

			String targetPathCount = folderPath + "\\Count_" + fileName;

			FileOutputStream outputStream11 = new FileOutputStream(targetPathCount);
			workBookOutput1.write(outputStream11);

			workBookOutput1.close();
			workBookCount.close();

			JOptionPane.showMessageDialog(null, "Count Excel created", "Excel", JOptionPane.PLAIN_MESSAGE);
			System.out.println("Count1......Done");

		} catch (NullPointerException ne) {
			count++;
//			ne.printStackTrace();
		} catch (FileNotFoundException e1) {
//			e1.printStackTrace();
			count++;
		} catch (IOException ee) {
//			ee.printStackTrace();
			count++;
		}

		return count;
	}

	
	// systemField

	//1
	static int SystemFile1Sheet = 0;
	static int selectedCounted = 22;
	static String selectedCountedName = "Booking Type";
	static String filePath1Count = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles\\HCMEL report.xlsx";
	static String fileName1ForCount = "HCMEL report.xlsx";
	
	//2
//	static int SystemFile1Sheet = 4;	
//	static String filePath1Count = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles\\Fresher Master Data .xlsx";
//	static String fileName1ForCount = "Fresher Master Data .xlsx";
//	static int selectedCounted = 11;
//	static String selectedCountedName = "Client Group Name";
	
	
	static String targetFolderForCount = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles";

	

	static Desktop desktop = Desktop.getDesktop();

	public static void main(String[] args) {

		int a = countExcel(filePath1Count, targetFolderForCount, fileName1ForCount, selectedCounted,
				selectedCountedName, SystemFile1Sheet);

		if (a < 1) {
			int ii = JOptionPane.showConfirmDialog(null,
					"We Have to close this window in order to open newly generated Excel, Because these are already open or are in use by javaw.exe Or if have to get more excels then click on No",
					"Exit?", JOptionPane.YES_NO_OPTION);
			if (ii == 1) {
				// do nothing
			}
			if (ii == 0) {
				try {
					File file = new File(targetFolderForCount);
					desktop.open(file);
				} catch (IOException eeee) {
					eeee.printStackTrace();
				}
				System.exit(0);
			}
		} else {
			JOptionPane.showMessageDialog(null, "Excels creation NOT DONE/File is missing - Or File is already opened!",
					"Excel !", JOptionPane.ERROR_MESSAGE);
		}
	}

}
