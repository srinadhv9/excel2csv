package excel2csv;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel2csv {

	public static void main(String[] args) {
		try{
		// reading file from desktop
		File inputFile = new File("C://Users//502362723//Desktop//14-Mar-2017 IT CPS VTW with WU.xlsm");
		// writing excel data to csv
		File outputFile = new File("C://Users//502362723//Desktop//excel2csv.csv");
		convertToXlsx(inputFile, outputFile);
		}catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public static void convertToXlsx(File inputFile, File outputFile) {
		StringBuffer bf = new StringBuffer();
		FileOutputStream fos = null;
		String strGetValue = "";
		XSSFWorkbook wb = null;
		try {
			fos = new FileOutputStream(outputFile);
			wb = new XSSFWorkbook(new FileInputStream(inputFile));
			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				System.out.println(wb.getSheetAt(i).getSheetName());
				if (wb.getSheetAt(i).getSheetName().equalsIgnoreCase("Raw Data")) {

					XSSFSheet sheet = wb.getSheetAt(i);
					Row row;
					Cell cell;
					//int intRowCounter = 0;
					Iterator<Row> rowIterator = sheet.iterator();
					while (rowIterator.hasNext()) {
						StringBuffer cellDData = new StringBuffer();
						row = rowIterator.next();
						int maxNumOfCells = sheet.getRow(0).getLastCellNum();
						int cellCounter = 0;
						while ((cellCounter) < maxNumOfCells) {
							if (sheet.getRow(row.getRowNum()) != null
									&& sheet.getRow(row.getRowNum()).getCell(cellCounter) != null) {
								cell = sheet.getRow(row.getRowNum()).getCell(cellCounter);
								CellValue cellValue = evaluator.evaluate(cell);
								//System.out.println(cellValue);
								//System.out.println(cellValue.getCellType());
								int cel_Type;
								try {
								cel_Type = cellValue.getCellType();
								} catch (NullPointerException e) {
								cel_Type = 3;
								}
								switch (cel_Type) {
								case Cell.CELL_TYPE_BOOLEAN:
									strGetValue = cellValue.getBooleanValue() + ",";
									cellDData.append(removeSpace(strGetValue));
									break;
								case Cell.CELL_TYPE_NUMERIC:
									strGetValue = new BigDecimal(cellValue.getNumberValue()).toPlainString();
									String tempStrGetValue = removeSpace(strGetValue);
									if (tempStrGetValue.length() == 0) {
										strGetValue = " ,";
										cellDData.append(strGetValue);
									} else {
										strGetValue = strGetValue + ",";
										cellDData.append(removeSpace(strGetValue));
									}
									break;
								case Cell.CELL_TYPE_STRING:
									strGetValue = cellValue.getStringValue().replaceAll(",", "~");
									String tempStrGetValue1 = removeSpace(strGetValue);
									if (tempStrGetValue1.length() == 0) {
										strGetValue = " ,";
										cellDData.append(strGetValue);
									} else {
										strGetValue = strGetValue + ",";
										cellDData.append(removeSpace(strGetValue));
									}
									break;
								case Cell.CELL_TYPE_BLANK:
									strGetValue = "" + ",";
									cellDData.append(removeSpace(strGetValue));
									break;
								default:
									strGetValue = cellValue + ",";
									cellDData.append(removeSpace(strGetValue));
								}
							} else {
								strGetValue = " ,";
								cellDData.append(strGetValue);
							}
							cellCounter++;
						}
						String temp = cellDData.toString();
						if (temp.endsWith(",")) {
							temp = temp.substring(0, temp.lastIndexOf(","));
							cellDData = null;
							bf.append(temp.trim());
						}
						bf.append("\n");
					//	intRowCounter++;
					}
					fos.write(bf.toString().getBytes());
					fos.close();
				}
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			try {
				if (fos != null)
					fos.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
			try {
				if (wb != null)
					wb.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}

	}

	private static String removeSpace(String strString) {
		if (strString != null && !strString.equals("")) {
			return strString.trim();
		}
		return strString;
	}
	
}
