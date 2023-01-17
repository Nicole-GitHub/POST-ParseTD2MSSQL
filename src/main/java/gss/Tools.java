package gss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Tools {
	private static final String className = Tools.class.getName();

	/**
	 * 取得 Excel的Workbook
	 * 
	 * @param path
	 * @return
	 */
	public static Workbook getWorkbook(String path) {
		Workbook workbook = null;
		InputStream inputStream = null;
		try {
			File f = new File(path);
			inputStream = new FileInputStream(f);
			String aux = path.substring(path.lastIndexOf(".") + 1);
			if ("XLS".equalsIgnoreCase(aux)) {
				workbook = new HSSFWorkbook(inputStream);
			} else if ("XLSX".equalsIgnoreCase(aux)) {
				workbook = new XSSFWorkbook(inputStream);
			} else {
				throw new Exception("檔案格式錯誤");
			}

		} catch (Exception ex) {
			// 因output時需要用到，故不可寫在finally內
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				throw new RuntimeException(className + " getWorkbook Error: \n" + e);
			}

			throw new RuntimeException(className + " getWorkbook Error: \n" + ex);
		} finally {
			try {
				if (inputStream != null)
					inputStream.close();
			} catch (IOException e) {
				throw new RuntimeException(className + " getWorkbook Error: \n" + e);
			}
		}
		return workbook;
	}

	/**
	 * 取得 Excel的Sheet
	 * 
	 * @param path
	 * @return
	 */
	public static Sheet getSheet(String path,String sheetName) {
		return getWorkbook(path).getSheet(sheetName);
	}

	/**
	 * 寫出整理好的Excel檔案
	 * @param workbook
	 * @param excelVersion
	 * @param outputPath
	 * @param outputFileName
	 */
	public static void output(Workbook workbook, String excelVersion, String outputPath, String outputFileName) {
		OutputStream output = null;
		File f = null;
		
		try {
			f = new File(outputPath);
			if(!f.exists()) f.mkdirs();
			
			f = new File(outputPath + outputFileName);
			output = new FileOutputStream(f);
			workbook.write(output);
		} catch (Exception ex) {
			throw new RuntimeException (className + " output Error: \n" + ex);
		} finally {
			try {
				if (workbook != null)
					workbook.close();
				if (output != null)
					output.close();
			} catch (IOException ex) {
				throw new RuntimeException (className + " output finally Error: \n" + ex);
			}
		}
	}

	/**
	 * 設定寫出檔案時的Style
	 */
	protected static CellStyle setStyleNormal(Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		short BorderStyle = CellStyle.BORDER_THIN;
		style.setBorderBottom(BorderStyle); // 儲存格格線(下)
		style.setBorderLeft(BorderStyle); // 儲存格格線(左)
		style.setBorderRight(BorderStyle); // 儲存格格線(右)
		style.setBorderTop(BorderStyle); // 儲存格格線(上)
		return style;
	}

	/**
	 * 設定寫出檔案時的Style
	 */
	protected static CellStyle setStyleError(Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(HSSFColor.RED.index);//儲存格底色為:紅色
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		short BorderStyle = CellStyle.BORDER_THIN;
		style.setBorderBottom(BorderStyle); // 儲存格格線(下)
		style.setBorderLeft(BorderStyle); // 儲存格格線(左)
		style.setBorderRight(BorderStyle); // 儲存格格線(右)
		style.setBorderTop(BorderStyle); // 儲存格格線(上)
		return style;
	}
	
	/**
	 * 設定Cell內容(含Style)
	 * 
	 * @param cell
	 * @param row
	 * @param cellNum
	 * @param cellValue
	 */
	public static void setCell(CellStyle style, Row row, int cellNum, String cellValue) {
		Cell cell = row.createCell(cellNum);
		cell.setCellValue(cellValue);
		cell.setCellStyle(style);
	}
	
	/**
     * 不為空
     */
	protected static boolean isntBlank(Cell cell) {
		return cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK;
	}
	
	/**
	 * 取今日日期 YYYY/MM/DD
	 * 
	 * @return
	 */
	public static String getToDay() {
		return new SimpleDateFormat("yyyy/MM/dd").format(new Date());
	}
}
