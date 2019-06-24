package com.sample.ExceltoMongo;

import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.LinkedHashMap;
import java.util.concurrent.ConcurrentHashMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;

public class Xls_Reader {
	public static String filename = "";
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private HSSFWorkbook workbook = null;
	private HSSFSheet sheet = null;
	private HSSFRow row = null;
	private HSSFCell cell = null;
	private HSSFRow header = null;
	DataFormatter fmt = new DataFormatter();
	static int count = 0;
	
	//public static ConcurrentHashMap<String, Object> workbookMap = new ConcurrentHashMap<String, Object>();
	public Xls_Reader(String path,ConcurrentHashMap<String, Object> workbookMap) throws Exception {
		this.path = path;
		fis = new FileInputStream(this.path);
		POIFSFileSystem fs = new POIFSFileSystem(fis);
		workbook = new HSSFWorkbook(fs);
		count =0;
		String name = null ;
		// sheet = workbook.getSheetAt(0);
		ConcurrentHashMap<String, Object> sheetMap = new ConcurrentHashMap<String, Object>();
		int i = workbook.getNumberOfSheets();
		for(int j =0 ;j <i;j++) {
		name = workbook.getSheetName(j);
		}
		makeSheetMap(sheetMap,name);
		try {
			if(workbook.getNumberOfSheets() == 4) {
			makeSheetMap(sheetMap, "Request Values");
			makeSheetMap(sheetMap, "Request Schema");
			makeSheetMap(sheetMap, "Response Values");
			makeSheetMap(sheetMap, "Response Schema");
			}
		} catch (Exception e) {
			//Log.info(path);
			e.printStackTrace();
		}
		try {
		if ((4 == count)) {
			workbookMap.put(path.substring((path.lastIndexOf("/") + 1), path.length()), sheetMap);
			//System.out.println(workbookMap);
		}
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		workbookMap.put(path.substring((path.lastIndexOf("/") + 1), path.length()), sheetMap);
		
		fis.close();
	}

	private void makeSheetMap(ConcurrentHashMap<String, Object> sheetMap, String sheetName) {
		try {
		HSSFSheet hssfSheet = workbook.getSheet(sheetName);
		if(hssfSheet!=null) {
		ConcurrentHashMap<Integer, Object> rowMap;
		ConcurrentHashMap<Integer, Object> cellMap;
		int physicalNumberOfRows;
		rowMap  = new ConcurrentHashMap<Integer, Object>();
		physicalNumberOfRows= getRowCount(sheetName)-1;
		int maxColumnCount = 0;
		
			HSSFRow row2 = hssfSheet.getRow(0);
			cellMap = new ConcurrentHashMap<Integer, Object>();
			try {
			for(int j=0;j<=row2.getPhysicalNumberOfCells();j++) {
				Object cellData = getCellDataForMap(sheetName, j, 0);
				if(cellData.equals("")) {
					maxColumnCount=j;
					break;
				}
				cellMap.put(j, cellData);
			}
			}
			catch (Exception e) {
			}
			rowMap.put(0, cellMap);
		
		for(int i=1;i<=physicalNumberOfRows;i++) {
			row2 = hssfSheet.getRow(i);
			boolean found = true;
			cellMap = new ConcurrentHashMap<Integer, Object>();
			for(int j=0;j<=maxColumnCount;j++) {
				Object cellData = getCellDataForMap(sheetName, j, i);
				if(getCellDataForMap(sheetName, 0, i).equals("")) {
					found = false;
					physicalNumberOfRows = i;
					break;
				}
				cellMap.put(j, cellData);
			}
			if(found == true) {
			rowMap.put(i, cellMap);
			}
		}
		sheetMap.put(sheetName, rowMap);
		count ++;
		}
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
		

	// returns the row count in a sheet
	public int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			sheet = workbook.getSheetAt(index);
			int number = sheet.getLastRowNum() + 1;
			return number;
		}
	}

	// returns the sheet count in the workbook
	public int getSheetCount() {
		int numberOfSheets = workbook.getNumberOfSheets();
		if (numberOfSheets == -1)
			return 0;
		else {
			return numberOfSheets;
		}
	}

	// returns the sheet name at the given index in the workbook
	public String getSheetName(int sheetIndex) {
		if (sheetIndex > -1) {
			String sheetName = workbook.getSheetName(sheetIndex);
			return sheetName;
		} else {
			return "Invalid Index";
		}
	}

	// returns the data from a cell
	public String getCellData(int sheetNo, int colNum, int rowNum) {
		try {
			if (rowNum <= 0)
				return "";
			sheet = workbook.getSheetAt(sheetNo);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(colNum);
			if (cell == null)
				return "";
			if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				// String cellText = String.valueOf(cell.getNumericCellValue());
				String cellText = fmt.formatCellValue(cell);
				if (!cellText.trim().isEmpty()) {
					// System.out.println("" + cellText);
				}
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();
					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.MONTH) + 1 + "/" + cal.get(Calendar.DAY_OF_MONTH) + "/" + cellText;
					// System.out.println(cellText);
				}
				return cellText;
			} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
			// else
			/*
			 * { String value = fmt.formatCellValue(cell); if (!
			 * value.trim().isEmpty()) { return
			 * String.valueOf(cell.getNumericCellValue()); //
			 * System.out.println("" + value); } }
			 */
		} catch (Exception e) {
			System.err.println("Exp - caused by : " + e.getCause());
			// e.printStackTrace();
			return "row " + rowNum + " or column " + colNum + " does not exist  in xls";
		}
	}

	// returns the data from a cell
	public String getCellData(String sheetName, String colName, int rowNum) {
		try {
			if (rowNum <= 0)
				return "";
			int index = workbook.getSheetIndex(sheetName);
			int col_Num = -1;
			if (index == -1)
				return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				try {
					if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
						col_Num = i;

				} catch(Exception e) {
					e.printStackTrace();
				}
				
			}
			if (col_Num == -1)
				return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(col_Num);
			if (cell == null)
				return "";
			// System.out.println(cell.getCellType());
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                return cell.getStringCellValue();

			} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				// String cellText = String.valueOf(cell.getStringCellValue());
				double numericCellValue = cell.getNumericCellValue();
				// int intValue = (int) numericCellValue;
				long longValue = (long) numericCellValue;
				String cellText = null;
				if (numericCellValue > (long) longValue) {
					cellText = String.valueOf(cell.getNumericCellValue());
				} else {
					cellText = String.valueOf(longValue);
				}
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();
					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR)));
					String month = String.valueOf((cal.get(Calendar.MONTH) + 1));
					int length = (int) (Math.log10(Integer.parseInt(month)) + 1);
					if (length < 2) {
						month = ("0" + month).toString();
					}
					String date = String.valueOf(cal.get(Calendar.DAY_OF_MONTH));
					int dateLength = (int) (Math.log10(Integer.parseInt(date)) + 1);
					if (dateLength < 2) {
						date = ("0" + date).toString();
					}
					cellText = cellText + "/" + (month) + "/" + date;
					// System.out.println(cellText);
				}
				return cellText;
			} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
		} catch (Exception e) {
			e.printStackTrace();
			return "row " + rowNum + " or column " + colName + " does not exist in xls";
		}
	}

	// returns the data from a cell
	public Object getCellData(String sheetName, int colNum, int rowNum) {
		try {
			if (rowNum <= 0)
				return "";
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(colNum);
			if (cell == null)
				return "";
			// System.out.println(cell.getCellType());
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				// String cellText = cell.getStringCellValue();
				// /commented on April 16 2015 as it was throwing exception when
				// formula is entered in a cell//
				double numericCellValue = cell.getNumericCellValue();
				// int intValue = (int) numericCellValue;
				long longValue = (long) numericCellValue;
				Object cellText = longValue;
				// if (numericCellValue > (long) longValue) {
				// cellText = String.valueOf(cell.getNumericCellValue());
				// } else {
				// cellText = String.valueOf(longValue);
				// }
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();
					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.MONTH) + 1 + "/" + cal.get(Calendar.DAY_OF_MONTH) + "/" + cellText;
					return cellText;
					// System.out.println(cellText);
				}
				return cellText;
			}
			if (cell.getCellType() == Cell.CELL_TYPE_STRING || cell.getCellType() == Cell.CELL_TYPE_FORMULA)
				return cell.getStringCellValue();
			else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
		} catch (Exception e) {
			System.err.println("Exp - caused by : " + e.getCause());
			// e.printStackTrace();
			return "row " + rowNum + " or column " + colNum + " does not exist  in xls";
		}
	}
	
	
	// returns the data from a cell
		public Object getCellDataForMap(String sheetName, int colNum, int rowNum) {
			try {
				
				int index = workbook.getSheetIndex(sheetName);
				if (index == -1)
					return "";
				sheet = workbook.getSheetAt(index);
				row = sheet.getRow(rowNum);
				if (row == null)
					return "";
				cell = row.getCell(colNum);
				if (cell == null)
					return "";
				// System.out.println(cell.getCellType());
				if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
					// String cellText = cell.getStringCellValue();
					// /commented on April 16 2015 as it was throwing exception when
					// formula is entered in a cell//
					double numericCellValue = cell.getNumericCellValue();
					// int intValue = (int) numericCellValue;
					long longValue = (long) numericCellValue;
					Object cellText = longValue;
					 if (numericCellValue > (long) longValue) {
					 cellText = String.valueOf(cell.getNumericCellValue());
					 } else {
					 cellText = String.valueOf(longValue);
					 }
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						// format in form of M/D/YY
						double d = cell.getNumericCellValue();
						Calendar cal = Calendar.getInstance();
						cal.setTime(HSSFDateUtil.getJavaDate(d));
						cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
						cellText = cal.get(Calendar.MONTH) + 1 + "/" + cal.get(Calendar.DAY_OF_MONTH) + "/" + cellText;
						return cellText;
						// System.out.println(cellText);
					}
					return cellText;
				}
				if (cell.getCellType() == Cell.CELL_TYPE_STRING || cell.getCellType() == Cell.CELL_TYPE_FORMULA)
					return cell.getStringCellValue();
				else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
					return "";
				else
					return String.valueOf(cell.getBooleanCellValue());
			} catch (Exception e) {
				System.err.println("Exp - caused by : " + e.getCause() + cell + "at row " + rowNum + " column :" +colNum + "sheetName : " + sheetName);
				// e.printStackTrace();
				return "row " + rowNum + " or column " + colNum + " does not exist  in xls";
			}
		}

	// returns true if data is set successfully else false
	public boolean setCellData(String sheetName, String colName, int rowNum, String data) {
		System.out.println("*****************Inside setCellData******************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			if (rowNum <= 0)
				return false;
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;
			sheet = workbook.getSheetAt(index);
			// int itest = row.getLastCellNum();
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				// String sTest = row.getCell(i).getStringCellValue().trim();
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;
			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			// cell style
			// CellStyle cs = workbook.createCellStyle();
			// cs.setWrapText(true);
			// cell.setCellStyle(cs);
			cell.setCellValue(data);
			fis.close();
			Thread.sleep(200);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// returns true if data is set successfully else false
	public boolean setCellData(String sheetName, String colName, int rowNum, String data, String url) {
		// System.out.println("setCellData setCellData******************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			if (rowNum <= 0)
				return false;
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;
			sheet = workbook.getSheetAt(index);
			// System.out.println("A");
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;
			sheet.autoSizeColumn(colNum); // ashish
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			cell.setCellValue(data);
			HSSFCreationHelper createHelper = workbook.getCreationHelper();
			// cell style for hyperlinks
			// by default hypelrinks are blue and underlined
			CellStyle hlink_style = workbook.createCellStyle();
			HSSFFont hlink_font = workbook.createFont();
			hlink_font.setUnderline(XSSFFont.U_SINGLE);
			hlink_font.setColor(IndexedColors.BLUE.getIndex());
			hlink_style.setFont(hlink_font);
			// hlink_style.setWrapText(true);
			HSSFHyperlink link = createHelper.createHyperlink(XSSFHyperlink.LINK_FILE);
			link.setAddress(url);
			cell.setHyperlink(link);
			cell.setCellStyle(hlink_style);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// returns true if sheet is created successfully else false
	public boolean addSheet(String sheetname) {
		FileOutputStream fileOut;
		try {
			workbook.createSheet(sheetname);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// returns true if sheet is removed successfully else false if sheet does
	// not exist
	public boolean removeSheet(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return false;
		FileOutputStream fileOut;
		try {
			workbook.removeSheetAt(index);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// returns true if column is created successfully
	public boolean addColumn(String sheetName, String colName) {
		// System.out.println("**************Inside
		// addColumn*********************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return false;
			Font font = workbook.createFont();
			font.setColor(HSSFColor.WHITE.index);
			font.setFontHeightInPoints((short) 12);
			HSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			/* setting white border */
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			// style.setBorderColor(BorderSide.LEFT,new XSSFColor(Color.WHITE));
			// style.setBorderColor(BorderSide.RIGHT,new
			// XSSFColor(Color.WHITE));
			// style.setBorderColor(BorderSide.TOP,new XSSFColor(Color.WHITE));
			// style.setBorderColor(BorderSide.BOTTOM,new
			// XSSFColor(Color.WHITE));
			style.setFont(font);
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			if (row == null)
				row = sheet.createRow(0);
			// cell = row.getCell();
			// if (cell == null)
			// System.out.println(row.getLastCellNum());
			if (row.getLastCellNum() == -1)
				cell = row.createCell(0);
			else
				cell = row.createCell(row.getLastCellNum());
			cell.setCellValue(colName);
			cell.setCellStyle(style);
			// fis.close();
			Thread.sleep(200);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// returns true if column is created successfully
	public boolean addColumnDarkGrey(String sheetName, String colName) {
		// System.out.println("**************Inside
		// addColumnDarkGrey*********************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return false;
			Font font = workbook.createFont();
			font.setColor(HSSFColor.WHITE.index);
			font.setFontHeightInPoints((short) 12);
			HSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.GREY_80_PERCENT.index);
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			/* setting the border white */
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			style.setFont(font);
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			if (row == null)
				row = sheet.createRow(0);
			// cell = row.getCell();
			// if (cell == null)
			// System.out.println(row.getLastCellNum());
			if (row.getLastCellNum() == -1)
				cell = row.createCell(0);
			else
				cell = row.createCell(row.getLastCellNum());
			cell.setCellValue(colName);
			cell.setCellStyle(style);
			fis.close();
			Thread.sleep(200);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// removes a column and all the contents
	public boolean removeColumn(String sheetName, int colNum) {
		try {
			if (!isSheetExist(sheetName))
				return false;
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			sheet = workbook.getSheet(sheetName);
			HSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
			// XSSFCreationHelper createHelper = workbook.getCreationHelper();
			style.setFillPattern(HSSFCellStyle.NO_FILL);
			for (int i = 0; i < getRowCount(sheetName); i++) {
				row = sheet.getRow(i);
				if (row != null) {
					cell = row.getCell(colNum);
					if (cell != null) {
						cell.setCellStyle(style);
						row.removeCell(cell);
					}
				}
			}
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// find whether sheets exists
	public boolean isSheetExist(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			index = workbook.getSheetIndex(sheetName.toUpperCase());
			if (index == -1)
				return false;
			else
				return true;
		} else
			return true;
	}

	// returns number of columns in a sheet
	public int getColumnCount(String sheetName) {
		// check if sheet exists
		if (!isSheetExist(sheetName))
			return -1;
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);
		if (row == null)
			return -1;
		return row.getLastCellNum();
	}

	// String sheetName, String testCaseName,String keyword ,String URL,String
	// message
	public boolean addHyperLink(String sheetName, String screenShotColName, String testCaseName, int index, String url,
			String message) {
		// System.out.println("ADDING addHyperLink******************");
		url = url.replace('\\', '/');
		if (!isSheetExist(sheetName))
			return false;
		sheet = workbook.getSheet(sheetName);
		for (int i = 2; i <= getRowCount(sheetName); i++) {
			if (getCellData(sheetName, 0, i).toString().equalsIgnoreCase(testCaseName)) {
				// System.out.println("**caught "+(i+index));
				setCellData(sheetName, screenShotColName, i + index, message, url);
				break;
			}
		}
		return true;
	}

	public int getCellRowNum(String sheetName, String colName, String cellValue) {
		for (int i = 2; i <= getRowCount(sheetName); i++) {
			if (getCellData(sheetName, colName, i).equalsIgnoreCase(cellValue)) {
				return i;
			}
		}
		return -1;
	}

	public int getCellRowNum(String sheetName, int colNum, String cellValue) {
		for (int i = 2; i <= getRowCount(sheetName); i++) {
			if (getCellData(sheetName, colNum, i).toString().equalsIgnoreCase(cellValue)) {
				return i;
			}
		}
		return -1;
	}

	public int getCellColNum(String sheetName, int rowNum, String cellValue) {
		for (int i = 0; i < getColumnCount(sheetName); i++) {
			if (getCellData(sheetName, i, rowNum).toString().equalsIgnoreCase(cellValue)) {
				return i;
			}
		}
		return -1;
	}

	/*****************************************************************/
	// /// Jan 31 2015
	// returns true if column is created successfully
	public boolean setColumnWidth(String sheetName, int colNum, int colWidth) {
		// System.out.println("**************Inside
		// setColumnWidth*********************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return false;
			if (colNum < 0)
				return false;
			sheet = workbook.getSheetAt(index);
			sheet.setColumnWidth(colNum, colWidth);
			return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
	}

	/**
	 * @param sheetName
	 * @param colName
	 * @return
	 */
	public boolean autoSizeCol(String sheetName, String colName) {
		// System.out.println("**************Inside
		// autoSizeCol******************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;
			sheet.autoSizeColumn(colNum);
			// fis.close();
			Thread.sleep(300);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	
	 * 
	
	 */
	// returns true if data is set successfully else false
	public boolean setCellDataWithColorAndCeterAlign(String sheetName, String colName, int rowNum, String data,
			String color) {
		// System.out.println("******************Inside
		// setCellDataWithColorAndCeterAlign*********************");
		// System.out.println("*******************************************************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			if (rowNum <= 0)
				return false;
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;
			// // sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			row.setHeight((short) 320);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			// cell style
			// CellStyle cs = workbook.createCellStyle();
			// cs.setWrapText(true);
			// cell.setCellStyle(cs);
			cell.setCellValue(data);
			// ///////////////////////
			// CellStyle style1 = resultWorkbook.createCellStyle();
			HSSFCellStyle style1 = workbook.createCellStyle();
			if (color.toLowerCase().contains("yellow"))
				style1.setFillForegroundColor(HSSFColor.YELLOW.index);
			else if (color.toLowerCase().contains("green"))
				style1.setFillForegroundColor(HSSFColor.GREEN.index);
			else if (color.toLowerCase().contains("blue"))
				style1.setFillForegroundColor(HSSFColor.BLUE.index);
			else if (color.toLowerCase().contains("orange"))
				style1.setFillForegroundColor(HSSFColor.ORANGE.index);
			else if (color.toLowerCase().contains("rose"))
				style1.setFillForegroundColor(HSSFColor.ROSE.index);
			else if (color.toLowerCase().contains("grey50"))
				style1.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
			else if (color.toLowerCase().contains("grey25"))
				style1.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
			else if (color.toLowerCase().contains("lightgreen"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
			else if (color.toLowerCase().contains("aqua"))
				style1.setFillForegroundColor(HSSFColor.AQUA.index);
			else if (color.toLowerCase().contains("indigo"))
				style1.setFillForegroundColor(HSSFColor.INDIGO.index);
			else if (color.toLowerCase().contains("automatic"))
				style1.setFillForegroundColor(HSSFColor.AUTOMATIC.index);
			else if (color.toLowerCase().contains("turquoise"))
				style1.setFillForegroundColor(HSSFColor.TURQUOISE.index);
			else if (color.toLowerCase().contains("tan"))
				style1.setFillForegroundColor(HSSFColor.TAN.index);
			else if (color.toLowerCase().contains("teal"))
				style1.setFillForegroundColor(HSSFColor.TEAL.index);
			else if (color.toLowerCase().contains("bluegrey"))
				style1.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
			else if (color.toLowerCase().contains("brightgreen"))
				style1.setFillForegroundColor(HSSFColor.BRIGHT_GREEN.index);
			else if (color.toLowerCase().contains("brown"))
				style1.setFillForegroundColor(HSSFColor.BROWN.index);
			else if (color.toLowerCase().contains("coral"))
				style1.setFillForegroundColor(HSSFColor.CORAL.index);
			else if (color.toLowerCase().contains("cornflowerblue"))
				style1.setFillForegroundColor(HSSFColor.CORNFLOWER_BLUE.index);
			else if (color.toLowerCase().contains("gold"))
				style1.setFillForegroundColor(HSSFColor.GOLD.index);
			else if (color.toLowerCase().contains("lavendar"))
				style1.setFillForegroundColor(HSSFColor.LAVENDER.index);
			else if (color.toLowerCase().contains("lemonchiffon"))
				style1.setFillForegroundColor(HSSFColor.LEMON_CHIFFON.index);
			else if (color.toLowerCase().contains("lightblue"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
			else if (color.toLowerCase().contains("lightcornflowerblue"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
			else if (color.toLowerCase().contains("lightorange"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
			else if (color.toLowerCase().contains("lightturquoise"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
			else if (color.toLowerCase().contains("lightyellow"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
			else if (color.toLowerCase().contains("lime"))
				style1.setFillForegroundColor(HSSFColor.LIME.index);
			else if (color.toLowerCase().contains("maroon"))
				style1.setFillForegroundColor(HSSFColor.MAROON.index);
			else if (color.toLowerCase().contains("olivegreen"))
				style1.setFillForegroundColor(HSSFColor.OLIVE_GREEN.index);
			else if (color.toLowerCase().contains("orchid"))
				style1.setFillForegroundColor(HSSFColor.ORCHID.index);
			else if (color.toLowerCase().contains("paleblue"))
				style1.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
			else if (color.toLowerCase().contains("pink"))
				style1.setFillForegroundColor(HSSFColor.PINK.index);
			else if (color.toLowerCase().contains("plum"))
				style1.setFillForegroundColor(HSSFColor.PLUM.index);
			else if (color.toLowerCase().contains("red"))
				style1.setFillForegroundColor(HSSFColor.RED.index);
			else if (color.toLowerCase().contains("royalblue"))
				style1.setFillForegroundColor(HSSFColor.ROYAL_BLUE.index);
			else if (color.toLowerCase().contains("seagreen"))
				style1.setFillForegroundColor(HSSFColor.SEA_GREEN.index);
			else if (color.toLowerCase().contains("skyblue"))
				style1.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
			else if (color.toLowerCase().contains("violet"))
				style1.setFillForegroundColor(HSSFColor.VIOLET.index);
			style1.setAlignment(HSSFCellStyle.VERTICAL_CENTER);
			style1.setVerticalAlignment(CellStyle.VERTICAL_TOP);
			style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
			/* setting bleck border */
			style1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			style1.setBorderRight(XSSFCellStyle.BORDER_THIN);
			style1.setBorderTop(XSSFCellStyle.BORDER_THIN);
			style1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			// style1.setBorderColor(BorderSide.LEFT,new
			// XSSFColor(Color.BLACK));
			// style1.setBorderColor(BorderSide.RIGHT,new
			// XSSFColor(Color.BLACK));
			// style1.setBorderColor(BorderSide.TOP,new XSSFColor(Color.BLACK));
			// style1.setBorderColor(BorderSide.BOTTOM,new
			// XSSFColor(Color.BLACK));
			// style1.setAlignment(HorizontalAlignment.CENTER);
			// XSSFCellStyle style = resultWorkbook.createCellStyle();
			// style1.setAlignment(0);(HorizontalAlignment.LEFT);
			style1.setWrapText(false);
			cell.setCellStyle(style1);
			// cell.setCellStyle(style);
			// ///////////////////////
			// fis.close();
			Thread.sleep(200);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	
	 * 
	
	 */
	// returns true if data is set successfully else false
	public boolean setCellDataWithColor(String sheetName, String colName, int rowNum, String data, String color) {
		// System.out.println("******************Inside
		// setCellDataWithColor********************");
		// System.out.println("*******************************************************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			if (rowNum <= 0)
				return false;
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;
			// // sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			row.setHeight((short) 320);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			cell.setCellValue(data);
			sheet.setDefaultRowHeightInPoints((float) 40.00);
			// ///////////////////////
			// CellStyle style1 = resultWorkbook.createCellStyle();
			HSSFCellStyle style1 = workbook.createCellStyle();
			if (color.toLowerCase().contains("yellow"))
				style1.setFillForegroundColor(HSSFColor.YELLOW.index);
			else if (color.toLowerCase().contains("lime"))
				style1.setFillForegroundColor(HSSFColor.LIME.index);
			else if (color.toLowerCase().contains("red"))
				style1.setFillForegroundColor(HSSFColor.RED.index);
			else if (color.toLowerCase().contains("lemonchiffon"))
				style1.setFillForegroundColor(HSSFColor.LEMON_CHIFFON.index);
			else if (color.toLowerCase().contains("coral"))
				style1.setFillForegroundColor(HSSFColor.CORAL.index);
			else if (color.toLowerCase().contains("gold"))
				style1.setFillForegroundColor(HSSFColor.GOLD.index);
			else if (color.toLowerCase().contains("green"))
				style1.setFillForegroundColor(HSSFColor.GREEN.index);
			else if (color.toLowerCase().contains("blue"))
				style1.setFillForegroundColor(HSSFColor.BLUE.index);
			else if (color.toLowerCase().contains("orange"))
				style1.setFillForegroundColor(HSSFColor.ORANGE.index);
			else if (color.toLowerCase().contains("rose"))
				style1.setFillForegroundColor(HSSFColor.ROSE.index);
			else if (color.toLowerCase().contains("grey50"))
				style1.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
			else if (color.toLowerCase().contains("grey25"))
				style1.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
			else if (color.toLowerCase().contains("white"))
				style1.setFillForegroundColor(HSSFColor.WHITE.index);
			else if (color.toLowerCase().contains("lightgreen"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
			else if (color.toLowerCase().contains("aqua"))
				style1.setFillForegroundColor(HSSFColor.AQUA.index);
			else if (color.toLowerCase().contains("indigo"))
				style1.setFillForegroundColor(HSSFColor.INDIGO.index);
			else if (color.toLowerCase().contains("automatic"))
				style1.setFillForegroundColor(HSSFColor.AUTOMATIC.index);
			else if (color.toLowerCase().contains("turquoise"))
				style1.setFillForegroundColor(HSSFColor.TURQUOISE.index);
			else if (color.toLowerCase().contains("tan"))
				style1.setFillForegroundColor(HSSFColor.TAN.index);
			else if (color.toLowerCase().contains("teal"))
				style1.setFillForegroundColor(HSSFColor.TEAL.index);
			else if (color.toLowerCase().contains("bluegrey"))
				style1.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
			else if (color.toLowerCase().contains("brightgreen"))
				style1.setFillForegroundColor(HSSFColor.BRIGHT_GREEN.index);
			else if (color.toLowerCase().contains("brown"))
				style1.setFillForegroundColor(HSSFColor.BROWN.index);
			else if (color.toLowerCase().contains("cornflowerblue"))
				style1.setFillForegroundColor(HSSFColor.CORNFLOWER_BLUE.index);
			else if (color.toLowerCase().contains("lavendar"))
				style1.setFillForegroundColor(HSSFColor.LAVENDER.index);
			else if (color.toLowerCase().contains("lightblue"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
			else if (color.toLowerCase().contains("lightcornflowerblue"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
			else if (color.toLowerCase().contains("lightorange"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
			else if (color.toLowerCase().contains("lightturquoise"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
			else if (color.toLowerCase().contains("lightyellow"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
			else if (color.toLowerCase().contains("maroon"))
				style1.setFillForegroundColor(HSSFColor.MAROON.index);
			else if (color.toLowerCase().contains("olivegreen"))
				style1.setFillForegroundColor(HSSFColor.OLIVE_GREEN.index);
			else if (color.toLowerCase().contains("orchid"))
				style1.setFillForegroundColor(HSSFColor.ORCHID.index);
			else if (color.toLowerCase().contains("paleblue"))
				style1.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
			else if (color.toLowerCase().contains("pink"))
				style1.setFillForegroundColor(HSSFColor.PINK.index);
			else if (color.toLowerCase().contains("plum"))
				style1.setFillForegroundColor(HSSFColor.PLUM.index);
			else if (color.toLowerCase().contains("royalblue"))
				style1.setFillForegroundColor(HSSFColor.ROYAL_BLUE.index);
			else if (color.toLowerCase().contains("seagreen"))
				style1.setFillForegroundColor(HSSFColor.SEA_GREEN.index);
			else if (color.toLowerCase().contains("skyblue"))
				style1.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
			else if (color.toLowerCase().contains("violet"))
				style1.setFillForegroundColor(HSSFColor.VIOLET.index);
			style1.setAlignment(HSSFCellStyle.VERTICAL_CENTER);
			style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
			/* Setting black Border */
			style1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			style1.setBorderRight(XSSFCellStyle.BORDER_THIN);
			style1.setBorderTop(XSSFCellStyle.BORDER_THIN);
			style1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			// style1.setBorderColor(BorderSide.LEFT,new
			// XSSFColor(Color.BLACK));
			// style1.setBorderColor(BorderSide.RIGHT,new
			// XSSFColor(Color.BLACK));
			// style1.setBorderColor(BorderSide.TOP,new XSSFColor(Color.BLACK));
			// style1.setBorderColor(BorderSide.BOTTOM,new
			// XSSFColor(Color.BLACK));
			style1.setWrapText(false);
			// style1.setWrapText(true);
			// style1.setAlignment(HorizontalAlignment.CENTER);
			// XSSFCellStyle style = resultWorkbook.createCellStyle();
			// style1.setAlignment(HorizontalAlignment.CENTER);
			cell.setCellStyle(style1);
			// cell.setCellStyle(style);
			// ///////////////////////
			// fis.close();
			Thread.sleep(200);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	
	 * 
	
	 */
	// returns true if data is set successfully else false
	public boolean setCellDataWithColorAndWrapText(String sheetName, String colName, int rowNum, String data,
			String color) {
		// System.out.println("******************Inside
		// setCellDataWithColorAndWrapText********************");
		System.out.println("*******************************************************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			if (rowNum <= 0)
				return false;
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				// System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;
			// // sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);
			cell.setCellValue(data);
			// ///////////////////////
			// CellStyle style1 = resultWorkbook.createCellStyle();
			HSSFCellStyle style1 = workbook.createCellStyle();
			if (color.toLowerCase().contains("yellow"))
				style1.setFillForegroundColor(HSSFColor.YELLOW.index);
			else if (color.toLowerCase().contains("lime"))
				style1.setFillForegroundColor(HSSFColor.LIME.index);
			else if (color.toLowerCase().contains("red"))
				style1.setFillForegroundColor(HSSFColor.RED.index);
			else if (color.toLowerCase().contains("lemonchiffon"))
				style1.setFillForegroundColor(HSSFColor.LEMON_CHIFFON.index);
			else if (color.toLowerCase().contains("coral"))
				style1.setFillForegroundColor(HSSFColor.CORAL.index);
			else if (color.toLowerCase().contains("gold"))
				style1.setFillForegroundColor(HSSFColor.GOLD.index);
			else if (color.toLowerCase().contains("green"))
				style1.setFillForegroundColor(HSSFColor.GREEN.index);
			else if (color.toLowerCase().contains("blue"))
				style1.setFillForegroundColor(HSSFColor.BLUE.index);
			else if (color.toLowerCase().contains("orange"))
				style1.setFillForegroundColor(HSSFColor.ORANGE.index);
			else if (color.toLowerCase().contains("rose"))
				style1.setFillForegroundColor(HSSFColor.ROSE.index);
			else if (color.toLowerCase().contains("grey50"))
				style1.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
			else if (color.toLowerCase().contains("grey25"))
				style1.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
			else if (color.toLowerCase().contains("white"))
				style1.setFillForegroundColor(HSSFColor.WHITE.index);
			else if (color.toLowerCase().contains("lightgreen"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
			else if (color.toLowerCase().contains("aqua"))
				style1.setFillForegroundColor(HSSFColor.AQUA.index);
			else if (color.toLowerCase().contains("indigo"))
				style1.setFillForegroundColor(HSSFColor.INDIGO.index);
			else if (color.toLowerCase().contains("automatic"))
				style1.setFillForegroundColor(HSSFColor.AUTOMATIC.index);
			else if (color.toLowerCase().contains("turquoise"))
				style1.setFillForegroundColor(HSSFColor.TURQUOISE.index);
			else if (color.toLowerCase().contains("tan"))
				style1.setFillForegroundColor(HSSFColor.TAN.index);
			else if (color.toLowerCase().contains("teal"))
				style1.setFillForegroundColor(HSSFColor.TEAL.index);
			else if (color.toLowerCase().contains("bluegrey"))
				style1.setFillForegroundColor(HSSFColor.BLUE_GREY.index);
			else if (color.toLowerCase().contains("brightgreen"))
				style1.setFillForegroundColor(HSSFColor.BRIGHT_GREEN.index);
			else if (color.toLowerCase().contains("brown"))
				style1.setFillForegroundColor(HSSFColor.BROWN.index);
			else if (color.toLowerCase().contains("cornflowerblue"))
				style1.setFillForegroundColor(HSSFColor.CORNFLOWER_BLUE.index);
			else if (color.toLowerCase().contains("lavendar"))
				style1.setFillForegroundColor(HSSFColor.LAVENDER.index);
			else if (color.toLowerCase().contains("lightblue"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
			else if (color.toLowerCase().contains("lightcornflowerblue"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
			else if (color.toLowerCase().contains("lightorange"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
			else if (color.toLowerCase().contains("lightturquoise"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
			else if (color.toLowerCase().contains("lightyellow"))
				style1.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
			else if (color.toLowerCase().contains("maroon"))
				style1.setFillForegroundColor(HSSFColor.MAROON.index);
			else if (color.toLowerCase().contains("olivegreen"))
				style1.setFillForegroundColor(HSSFColor.OLIVE_GREEN.index);
			else if (color.toLowerCase().contains("orchid"))
				style1.setFillForegroundColor(HSSFColor.ORCHID.index);
			else if (color.toLowerCase().contains("paleblue"))
				style1.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
			else if (color.toLowerCase().contains("pink"))
				style1.setFillForegroundColor(HSSFColor.PINK.index);
			else if (color.toLowerCase().contains("plum"))
				style1.setFillForegroundColor(HSSFColor.PLUM.index);
			else if (color.toLowerCase().contains("royalblue"))
				style1.setFillForegroundColor(HSSFColor.ROYAL_BLUE.index);
			else if (color.toLowerCase().contains("seagreen"))
				style1.setFillForegroundColor(HSSFColor.SEA_GREEN.index);
			else if (color.toLowerCase().contains("skyblue"))
				style1.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
			else if (color.toLowerCase().contains("violet"))
				style1.setFillForegroundColor(HSSFColor.VIOLET.index);
			HSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
			style1.setAlignment(HSSFCellStyle.VERTICAL_CENTER);
			style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
			/* Setting black Border */
			style1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			style1.setBorderRight(XSSFCellStyle.BORDER_THIN);
			style1.setBorderTop(XSSFCellStyle.BORDER_THIN);
			style1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			style1.setWrapText(true);
			cell.setCellStyle(style1);
			// fis.close();
			Thread.sleep(200);
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/**
	
	 * 
	
	 */
	// returns true if column is created successfully
	public boolean isColumnExist(String sheetName, String colName) {
		// System.out.println("**************Inside
		// isColumnExist*********************");
		try {
			fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return false;
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			if (row == null)
				return false;
			// row = sheet.createRow(0);
			if (row.getLastCellNum() == -1)
				return false;
			// cell = row.createCell(0);
			// else
			// cell = row.createCell(row.getLastCellNum());
			int noOfCols = row.getLastCellNum();
			// System.out.println("nO OF COLS : "+noOfCols);
			// System.out.println(row.getCell(0).getStringCellValue().trim());
			// String cellVal="";
			for (int i = 2; i < noOfCols; i++) {
				// cellVal=row.getCell(i).getStringCellValue().trim();
				if (row.getCell(i).getStringCellValue().trim().equals(colName.trim())) {
					// if(cellVal.equalsIgnoreCase(colName.trim())) {
					System.out.println("Column Exist - " + row.getCell(i).getStringCellValue().toString().trim());
					// System.out.println("YA GOT IT AT - "+i);
					return true;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return false;
	}

	/*
	 * This method reads the instruction file column by column according to row
	 * no passed in this method and puts all the values of the column of the
	 * particular row in a map .
	 */
	public LinkedHashMap<String, String> getCellData(String sheetName, int rowNo) {
		// List<Object> headers=new ArrayList<>();
		// headers=null;
		LinkedHashMap<String, String> objects = new LinkedHashMap<String, String>();
		try {
			int index = workbook.getSheetIndex(sheetName);
			int col_Num = -1;
			if (index == -1)
				return objects;
			sheet = workbook.getSheetAt(index);
			// for(int i = 0; i < row.getLastCellNum(); i++) {
			// headers.add(row.getCell(i).getStringCellValue());
			//
			// }
			// int countRows=sheet.getLastRowNum();
			// for(int i=1;i<countRows;i++){
			row = sheet.getRow(rowNo - 1);
			if (row == null)
				return objects;
			for (int j = 0; j < row.getLastCellNum(); j++) {
				// sheet = workbook.getSheetAt(index);
				header = sheet.getRow(0);
				cell = row.getCell(j);
				if (cell == null)
					objects.put(header.getCell(j).getStringCellValue(), "");
				// System.out.println(cell.getCellType());
				else if (cell.getCellType() == Cell.CELL_TYPE_STRING)
					objects.put(header.getCell(j).getStringCellValue(), cell.getStringCellValue());
				else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
					// String cellText =
					// String.valueOf(cell.getStringCellValue());
					double numericCellValue = cell.getNumericCellValue();
					// int intValue = (int) numericCellValue;
					long longValue = (long) numericCellValue;
					String cellText = null;
					if (numericCellValue > (long) longValue) {
						cellText = String.valueOf(cell.getNumericCellValue());
					} else {
						cellText = String.valueOf(longValue);
					}
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						// format in form of M/D/YY
						double d = cell.getNumericCellValue();
						Calendar cal = Calendar.getInstance();
						cal.setTime(HSSFDateUtil.getJavaDate(d));
						cellText = (String.valueOf(cal.get(Calendar.YEAR)));
						String month = String.valueOf((cal.get(Calendar.MONTH) + 1));
						int length = (int) (Math.log10(Integer.parseInt(month)) + 1);
						if (length < 2) {
							month = ("0" + month).toString();
						}
						String date = String.valueOf(cal.get(Calendar.DAY_OF_MONTH));
						int dateLength = (int) (Math.log10(Integer.parseInt(date)) + 1);
						if (dateLength < 2) {
							date = ("0" + date).toString();
						}
						cellText = cellText + "/" + (month) + "/" + date;
						// System.out.println(cellText);
					}
					objects.put(header.getCell(j).getStringCellValue(), cellText);
				} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
					objects.put(header.getCell(j).getStringCellValue(), "");
				else
					objects.put(header.getCell(j).getStringCellValue(), String.valueOf(cell.getBooleanCellValue()));
			}
			// headers.add(objects);
			// objects.clear();
			// }
		} catch (Exception e) {
			e.printStackTrace();
		}
		return objects;
	}

	public void closeWorkBook() {
		workbook = null;
	}
}
