package com.swing2;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;

import javax.swing.JOptionPane;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DemoCount {

	private static int countExcel(String filePath, String folderPath, String fileName, int selectedSheet,
			String selectedCountedName, ArrayList<Integer> countArray) {

		System.out.println("inside countExcel");

		for (Integer integer : countArray) {
			System.out.print(integer + ",");
		}

		int count = 0;

		try {

			// reading purpose
			FileInputStream file1Count = new FileInputStream(filePath);
			XSSFWorkbook workBookCount = new XSSFWorkbook(file1Count);
			XSSFSheet sheetCount = workBookCount.getSheetAt(selectedSheet);

			int totalNumberOfRowsInExcelCount = sheetCount.getLastRowNum();

			for (int i = 0; i < countArray.size(); i++) {
				int columnIndex = countArray.get(i);
			
			
			int total = 0;
			double totalP = 0;

			Set<String> set = new HashSet<>();

			for (int r = 1; r <= totalNumberOfRowsInExcelCount; r++) {

				if (sheetCount.getRow(r) == null) {
					continue;
				}
				if (sheetCount.getRow(r).getCell(columnIndex) == null) {
					continue;
				}
				set.add(sheetCount.getRow(r).getCell(columnIndex).toString());
			}

			String[] setToStringArr = set.toArray(new String[set.size()]);
			int[] arr = new int[set.size()];

			for (int r = 0; r < setToStringArr.length; r++) {

				for (int j = 1; j <= totalNumberOfRowsInExcelCount; j++) {

					if (sheetCount.getRow(j) == null) {
						continue;
					}
					if (sheetCount.getRow(j).getCell(columnIndex) == null) {
						continue;
					}

					if (setToStringArr[r].equalsIgnoreCase(sheetCount.getRow(j).getCell(columnIndex).toString())) {
						arr[r]++;
					}
				}
			}

			for (int count1 : arr) {
//				System.out.println(count);
				total += count1;
//				totalP += count1;
			}

			// percentage
			double[] arrPer = new double[set.size()];
//			double percentage = 0;

			for (int count1 : arr) {
//				System.out.println(count1);
				totalP += count1;
			}

//			System.out.println("total:"+totalP);

			for (int ii = 0; ii < arrPer.length; ii++) {
				arrPer[ii] = (arr[ii] / totalP) * 100;
			}
			
			}
			
			
			

			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;

			for (int r = 0; r <= setToStringArr.length; r++) {
				rowCreated = sheetCreate1.createRow(r);

				for (int c = 0; c < 4; c++) {
					rowCreated.createCell(c);
				}
			}

			for (int c = 0; c < 4; c++) {
				for (int i = 0; i <= setToStringArr.length; i++) {

					if (i < setToStringArr.length) {

						if (c == 0 && i == 0) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(selectedCountedName);
						} else if (c == 1) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(setToStringArr[i]);
						} else if (c == 2) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(arr[i]);
						} else if (c == 3) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(String.format("%.2f", arrPer[i]) + " %");
						}
					}

					if (i <= setToStringArr.length) {

						if (c == 1 && i == setToStringArr.length) {
							sheetCreate1.getRow(i).getCell(c).setCellValue("total:");
						} else if (c == 2 && i == setToStringArr.length) {
							sheetCreate1.getRow(i).getCell(c).setCellValue(total);
						} else if (c == 3 && i == setToStringArr.length) {
							sheetCreate1.getRow(i).getCell(c).setCellValue("100%");
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
//			count++;
			ne.printStackTrace();
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
//			count++;
		} catch (IOException ee) {
			ee.printStackTrace();
//			count++;
		}
		return count;
	}

	public static void main(String[] args) {

		ArrayList<Integer> arr = new ArrayList<>();
		arr.add(2);
		arr.add(4);
		arr.add(5);
		arr.add(6);
		arr.add(7);
		
		countExcel("C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles\\Fresher Master Data .xlsx"
				, "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles", "Fresher Master Data .xlsx",
				3, null, arr);
	}

}
