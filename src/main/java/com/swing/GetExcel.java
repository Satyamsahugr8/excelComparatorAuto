package com.swing;

import java.io.FileOutputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetExcel {

	public static void main(String[] args) {

		try {

			//1
			XSSFWorkbook wb = new XSSFWorkbook(OPCPackage.open(
					"C:\\Users\\SATYASAH\\eclipse-Myworkspace\\"
							+ "CGExcelAuto\\Copy of RrecruitmentReport 08-05-23 cutoff time 9.00 am.xlsx",
					PackageAccess.READ));

			XSSFSheet sheet1 = wb.getSheetAt(0);
			
			int totalNumberOfRows1 = sheet1.getLastRowNum();
			int totalNumberOfColumn1 = sheet1.getRow(0).getLastCellNum();

			System.out.println("totalNumberOfRows1:" + totalNumberOfRows1);
			System.out.println("totalNumberOfColumn1:" + totalNumberOfColumn1);
			
			int keyFile1 = 1;
			
			
			//2
			XSSFWorkbook wb2 = new XSSFWorkbook(OPCPackage.open(
					"C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles\\same.xlsx",
					PackageAccess.READ));

			Sheet sheet2 = wb2.getSheetAt(0);
			
			int totalNumberOfRows2 = sheet2.getLastRowNum();
			int totalNumberOfColumn2 = sheet2.getRow(0).getLastCellNum();
			
			System.out.println("totalNumberOfRows2:" + totalNumberOfRows2);
			System.out.println("totalNumberOfColumn2:" + totalNumberOfColumn2);

			int keyFile2 = 1;
			

			// going to Excel1 (key column) row => 1 to last
			for (int r = 1; r <= totalNumberOfRows1; r++) {

				if (sheet1.getRow(r) == null) {
					continue;
				} else {

					int counterr = 0;

					if (sheet1.getRow(r).getCell(keyFile1) == null) {
						counterr = 0;
					} else {

						// going to Excel2 (key column) row => 1 to last
						for (int e = 1; e <= totalNumberOfRows2; e++) {

							if (sheet2.getRow(e) == null) {
								continue;
							} else {

								if (sheet2.getRow(e).getCell(keyFile2) == null) {
									continue;
								}

								if (sheet1.getRow(r).getCell(keyFile1).toString()
										.equals(sheet2.getRow(e).getCell(keyFile2).toString())) {
//												System.out.println("SameCells1:" + sheet1.getRow(r).getCell(keyFile1) + "/"+ sheet2.getRow(e).getCell(keyFile2));
									counterr++;
									break;
								}
							}
						}
					} // else

					if (counterr == 0) {
						XSSFRow row = sheet1.getRow(r);
						sheet1.removeRow(row);
					}

				} // else
			} // for
			
			
			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;
			
			// counting null row in EXCEL 1
						int counter = 0;
						for (int r = 1; r <= totalNumberOfRows1; r++) {
							if (sheet1.getRow(r) == null) {
								counter++;
							}
						}

						

						if (counter != 0) {

							// creating new excel 1 removing NULL row
							int totalNumberOfRowsOfNewSheet = totalNumberOfRows1 - counter;

							for (int r = 0; r <= totalNumberOfRowsOfNewSheet; r++) {
								rowCreated = sheetCreate1.createRow(r);

								for (int c = 0; c < totalNumberOfColumn1; c++) {
									rowCreated.createCell(c);
								}
							}

							for (int p = 0, u = 0; p <= totalNumberOfRows1; p++) {

								if (sheet1.getRow(p) == null) {
									continue;
								} else {

									rowCreated = sheetCreate1.getRow(u);

									for (int d = 0; d < totalNumberOfColumn1; d++) {

										if (sheet1.getRow(p).getCell(d) == null) {
											continue;
										} else {

											if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
												rowCreated.getCell(d)
														.setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
											} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
												rowCreated.getCell(d)
														.setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
											} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
												rowCreated.getCell(d)
														.setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
											}

										}
									}
								}
								u++;
							}
						}

//			// creating new working and adding new rows for excel1
//			XSSFWorkbook workBookOutput = new XSSFWorkbook();
//			XSSFSheet sheetCreate1 = workBookOutput.createSheet();
//			XSSFRow rowCreated = null;
//
//			for (int rr = 0; rr <= totalNumberOfRows; rr++) {
//				rowCreated = sheetCreate1.createRow(rr);
//
//				for (int c = 0; c < totalNumberOfColumn; c++) {
//					rowCreated.createCell(c);
//				}
//			}
//
//			for (int r = 0; r <= totalNumberOfRows; r++) {
//
//				for (int c = 0; c < totalNumberOfColumn; c++) {
//					if (sheet.getRow(r) == null) {
//						continue;
//					}
//					if (sheet.getRow(r).getCell(c) == null) {
//						continue;
//					} else {
//						rowCreated = sheetCreate1.getRow(r);
//
//						if (sheet.getRow(r).getCell(c).getCellType() == CellType.STRING) {
//							rowCreated.getCell(c).setCellValue(sheet.getRow(r).getCell(c).getStringCellValue());
//						} else if (sheet.getRow(r).getCell(c).getCellType() == CellType.NUMERIC) {
//							rowCreated.getCell(c).setCellValue(sheet.getRow(r).getCell(c).getNumericCellValue());
//						} else if (sheet.getRow(r).getCell(c).getCellType() == CellType.BOOLEAN) {
//							rowCreated.getCell(c).setCellValue(sheet.getRow(r).getCell(c).getBooleanCellValue());
//						}
//
//					}
////				System.out.println();
//				}
//			}

			System.out.println("Excel created");
			String target1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\Output\\Enter.xlsx";
			FileOutputStream outputStream22 = new FileOutputStream(target1Path);
			workBookOutput1.write(outputStream22);
			workBookOutput1.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
