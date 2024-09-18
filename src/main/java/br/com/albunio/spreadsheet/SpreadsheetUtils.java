package br.com.albunio.spreadsheet;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Jaime Albunio
 * @since April 10, 2020
 * @version 20240109
 */
public class SpreadsheetUtils {
	
	private static DataFormatter DATA_FORMATTER = null;
	
	/**
	 * Verify if cellRPlus3 is a number cell (NUMBER or NUMBER_FORMULA or number as String)
	 * @since April 10, 2020
	 * @version 20200422
	 * @param cellRPlus3
	 * @return
	 */
	public static boolean isNumber(Cell cellRPlus3) {
		if(cellRPlus3.getCellType().equals(CellType.NUMERIC) || cellRPlus3.getCellType().equals(CellType.FORMULA)) {
			return true;
		}
		if(cellRPlus3.getCellType().equals(CellType.STRING) && cellRPlus3.getStringCellValue() != null) {
			try {
				Float floatN = Float.parseFloat(cellRPlus3.getStringCellValue());
				if(floatN != null) {
					return true;
				}
			} catch (NullPointerException | NumberFormatException e) {
			}
		}
		return false;
	}
	
	/**
	 * 
	 * @since May 21, 2021
	 * @version 20210521
	 * @param cell
	 * @return
	 */
	public static String getNumberAsString(Cell cell) {
		String resultString = null;
		
		if(cell.getCellType().equals(CellType.NUMERIC)) {
			return Double.toString(cell.getNumericCellValue());
		}
		
		if(cell.getCellType().equals(CellType.FORMULA)) {
			
		}else if(cell.getCellType().equals(CellType.STRING) && cell.getStringCellValue() != null) {
			return cell.getStringCellValue();
		}
		
		return resultString;
	}

	/**
	 * 
	 * @since Jun 1, 2021
	 * @version 20210601
	 * @param cell
	 * @return
	 */
	public static String getCellAsString(Cell cell) {
		String resultString = null;
		
		if(cell != null) {
			
			if(cell.getCellType().equals(CellType.BLANK)) {
				return "";
			}
			
			if(cell.getCellType().equals(CellType.NUMERIC)) {
				return Double.toString(cell.getNumericCellValue());
			}
			
			if(cell.getCellType().equals(CellType.FORMULA)) {
				if(cell.getCachedFormulaResultType().equals(CellType.NUMERIC)) {
					return Double.toString(cell.getNumericCellValue());
				}
				if(cell.getCachedFormulaResultType().equals(CellType.STRING)) {
					return cell.getRichStringCellValue().getString();
				}
				if(cell.getCachedFormulaResultType().equals(CellType.ERROR)) {
					resultString = "ERROR";
				}
			}else if(cell.getCellType().equals(CellType.STRING) && cell.getStringCellValue() != null) {
				return cell.getStringCellValue();
			}
		}
		
		return resultString;
	}
	
	/**
	 * 
	 * @since April 20, 2020
	 * @version 20200420
	 * @param fileName
	 * @return
	 * @throws IOException
	 */
	public static Workbook getWorkbook(String fileName) throws IOException {
		FileInputStream stream = new FileInputStream(fileName);
		if(fileName.endsWith(".xlsx")) {
			XSSFWorkbook wb = new XSSFWorkbook(stream);
			return wb;
		}else {
			HSSFWorkbook wb = new HSSFWorkbook(stream);
			return wb;
		}
	}
	
	/**
	 * 
	 * @since April 28, 2020
	 * @version 20200428
	 * @return
	 */
	public static DataFormatter getDataFormatter() {
		if(DATA_FORMATTER == null) {
			DATA_FORMATTER = new DataFormatter(new Locale("pt", "BR"));
		}
		
		return DATA_FORMATTER;
	}
}