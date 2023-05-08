package gss;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Parser {
	private static final String className = Parser.class.getName();

	public static void main(String args[]) throws Exception {
		
		String os = System.getProperty("os.name");

		System.out.println("=== NOW TIME: " + new Date());
		System.out.println("=== os.name: " + os);
		String svnPath = "";
		
		// 判斷當前執行的啟動方式是IDE還是jar
		boolean isStartupFromJar = new File(Parser.class.getProtectionDomain().getCodeSource().getLocation().getPath()).isFile();
		System.out.println("=== isStartupFromJar: " + isStartupFromJar);
		
		// 程式執行完後欲產出的檔案Excel & SQL Path
//		String targetTableLayoutExcelPath = System.getProperty("user.dir") + File.separator; // Jar
		String path = System.getProperty("user.dir") + File.separator; // Jar
		if(!isStartupFromJar) {// IDE
			path = os.contains("Mac") ? "/Users/nicole/Dropbox/POST/JavaTools/POST-ParseTD2MSSQL/" // Mac
							: "C:/Users/nicole_tsou/Dropbox/POST/JavaTools/POST-ParseTD2MSSQL/"; // win
//			targetTableLayoutExcelPath = "C:/Users/nicole_tsou/Dropbox/POST/JavaTools/POST-ParseTD2MSSQL/";
			svnPath = "C:/22/DW/dw2209/";
		}
		
		/**
		 * 透過windows的cmd執行時需將System.in格式轉為big5才不會讓中文變亂碼
		 * 即使在cmd下chcp 65001轉成utf-8也沒用
		 * 但在eclipse執行時不能轉為big5
		 */
		Scanner s = null;
		try {
			s =  isStartupFromJar ? new Scanner(System.in, "big5") : new Scanner(System.in);
			System.out.println("請輸入您本機放置SVN(dw2209)目錄的路徑(例:C:/22/DW/dw2209/): ");
			svnPath = "".equals(svnPath) ? s.nextLine() + "/" : svnPath;
		}catch(Exception ex) {
		}finally {
			if(s != null) s.close();
		}
		
		// 要移轉的清單Excel與對應的SQL檔
		String sourceSQLListExcelPath = svnPath + "DOCUMENT/1-REQ/儲壽功能_檔案清單.xlsx";
		String sourceSQLPath = svnPath + "COLLECTION/郵政整體資訊管理系統/現行郵政整體資訊管理系統SourceCode/TableScript/Schema/";
		// 要與上述SQL比對的Table Spec
		String sourceTableLayoutExcelPath = svnPath + "DOCUMENT/3-SD/DW/Table Spec/";

		String output = path + "Output/";
		String targetSQLPath = output + "ParseTD2MSSQLSchema/";
		String targetTableLayoutExcelPath = output + "ParseTD2TableLayout/";
				
		Map<String,String> mapPath = new HashMap<String,String>();
		mapPath.put("sourceSQLListExcelPath", sourceSQLListExcelPath);
		mapPath.put("sourceSQLPath", sourceSQLPath);
		mapPath.put("sourceTableLayoutExcelPath", sourceTableLayoutExcelPath);
		mapPath.put("targetTableLayoutExcelPath", targetTableLayoutExcelPath);
		mapPath.put("targetSQLPath", targetSQLPath);
		

		FileTools.deleteFolder(path + "Output/");
		List<Map<String, String>> sourceTableList = runParserSourceSQLListExcel(Tools.getSheet(sourceSQLListExcelPath, "檔案清單(盤點)"));
		
		runParserSourceSQL(mapPath, sourceTableList);
	}
	
	/**
	 * 解悉出要移轉的清單
	 * @param sheetTable
	 * @return
	 * @throws Exception
	 */
	private static List<Map<String, String>> runParserSourceSQLListExcel(Sheet sheet) throws Exception {
		Row row = null;
		Cell cell = null;
		int spaceRow = 0;
		List<Map<String, String>> mapList = new ArrayList<Map<String, String>>();
		Map<String, String> map = new HashMap<String, String>();

		try {
			// 找出欲解析的資料有幾行
			for (int i = 0; i <= sheet.getLastRowNum(); i++) {
				// 因文件下方有很多多餘空白行，為不影響效能，若遇到超過連三行皆空白時表示資料行已到底，不需再讀
				if (spaceRow > 3)
					break;
				row = sheet.getRow(i);
				cell = row == null ? null : row.getCell(0);
				if (cell != null && "TABLE".equals(cell.toString().toUpperCase().trim())) {
					map = new HashMap<String, String>();
					map.put("TableName", row.getCell(4).toString());
					map.put("SubSys", row.getCell(11).toString());
					mapList.add(map);
					spaceRow = 0;
				} else if (cell == null || StringUtils.isBlank(cell.toString())) {
					spaceRow++;
				}
			}
		} catch (Exception ex) {
			throw new Exception(className + " Error: \n" + ex);
		}

System.out.println("runParserSourceSQLListExcel 檔案清單分析 Done! ");
		return mapList;
	}
	
	/**
	 * 1.解悉出要移轉的清單的Table Layout 
	 * 2.寫出移轉資料所需的Script
	 * @param path
	 * @param fileNameList
	 * @return
	 * @throws Exception
	 */
	private static void runParserSourceSQL(Map<String,String> mapPath, List<Map<String, String>> sourceTableList) throws Exception {
		String msDBName = "DDWQDHIS";
		String DATA_SOURCE = "SQLServerInstance";
		String createExternalTable = "",createMSSQLTable = "", insertInto = "";
		
		List<Map<String, String>> mapList = new ArrayList<Map<String, String>>();
		Map<String, String> map = new HashMap<String, String>();
		
		try {
			for (Map<String, String> mapTableInfo : sourceTableList) {
				// 取Teradata SQL Script 內容
				String sourceSQLPath = mapPath.get("sourceSQLPath")+ mapTableInfo.get("TableName") + ".sql";
//				System.out.println("sqlFileName:" + sourceSQLPath );
				String oriSQL = FileTools.readFileContent(sourceSQLPath);
				String tdDBTableNameDot = oriSQL.substring(oriSQL.toUpperCase().indexOf("TABLE") + 5,oriSQL.indexOf(",")).trim();
				String tdDBTableNameUL = tdDBTableNameDot.replace(".", "_");
				String dwDBTableName = msDBName + ".dbo." + tdDBTableNameUL;
				String polybaseTableName = "dbo.TD_" + tdDBTableNameUL;
//				if("DP_PMM.THPMSCR1".equals(tableName))
//					System.out.println("Stop!");
				
				// 截取schema部份
				String schema = oriSQL.substring(oriSQL.indexOf("("), oriSQL.lastIndexOf(")", oriSQL.indexOf("PRIMARY INDEX"))+1);
				schema = replaceDoubleSpace(schema.replace("CHARACTER SET LATIN NOT CASESPECIFIC", ""));
				schema = replaceDoubleSpace(schema.replace("CHARACTER SET UNICODE NOT CASESPECIFIC", ""));
				schema = replaceDoubleSpace(schema.replace("FORMAT 'YYYY/MM/DD'", ""));
				
				// 截取PRIMARY INDEX部份
				String pk = oriSQL.substring(oriSQL.indexOf("PRIMARY INDEX"));
				pk = pk.substring(pk.indexOf("(") + 1,pk.indexOf(")")).replace("\r\n", "").trim();
				
//				// 2.外部關聯表EXTERNAL TABLE
//				createExternalTable += "insert into BDBU_POST_HIS.dbo.CreateExternalTable \r\n"
//						+ "values('CREATE EXTERNAL TABLE " + tableName + "\r\n(" ;
//				// 2.CREATE MSSQL TABLE
//				createMSSQLTable += "insert into BDBU_POST_HIS.dbo.CreateMSSQLTable \r\n"
//						+ "values('CREATE TABLE " + dwTableName + "\r\n(" ;

				// 2.外部關聯表EXTERNAL TABLE
				createExternalTable += "CREATE EXTERNAL TABLE " + polybaseTableName + "\r\n(" ;
				// 2.CREATE MSSQL TABLE
				createMSSQLTable += "CREATE TABLE " + dwDBTableName + "\r\n(" ;
				// 2.insert into
//				insertInto += "insert into " + msDBName + ".dbo.InsertInto \r\n"
//						+ "values('insert into " + dwDBTableName + " select * from " + polybaseTableName + ";');\r\n";
				insertInto += "insert into " + dwDBTableName + " select * from " + polybaseTableName + ";\r\n";
				
				// 1.將上述的schema解悉為較細項的Table Layout
				mapList = new ArrayList<Map<String, String>>();
				String[] colSchemaList = schema.split("\r\n");
				String[] pkList = pk.split(",");
				for(String forColSchema : colSchemaList) {

					map = new HashMap<String, String>();
					String[] schemaList = forColSchema.split(" ");
					if (schemaList.length < 3)
						continue;
					
					String colName = schemaList[1].trim();
					String colTypeLen = schemaList[2].trim();
					String colType = colTypeLen.contains("(") 
							? colTypeLen.substring(0, colTypeLen.indexOf("(")).trim()
							: colTypeLen.replace(",", "");
					String colLen = colTypeLen.contains("(") 
							? colTypeLen.substring(colTypeLen.indexOf("(") + 1, colTypeLen.indexOf(")")).trim().replace(" ", "")
							: "";
					String colNull = forColSchema.contains("NOT NULL") ? "" : "Y";
					
					// PK
					String colPK = "";
					for (String str : pkList) {
						if(str.trim().equals(colName)) {
							colPK = "Y" ;
							break;
						}
					}
					
					// 將Teradata的欄位型態改為MSSQL的欄位型態
					if ("BYTEINT".equals(colType))
						colType = "SmallInt";
					else if ("FLOAT".equals(colType))
						colType = "Decimal";
					else if ("BYTE".equals(colType))
						colType = "Binary";
					else if ("VARBYTE".equals(colType))
						colType = "Varbinary";
					else if ("BLOB".equals(colType))
						colType = "varbinary";
					else if ("CHAR".equals(colType))
						colType = "Nchar";
					else if ("CLOB".equals(colType))
						colType = "nvarchar";
					else if ("VARCHAR".equals(colType))
						colType = "nvarchar";
					else if ("Graphic".equals(colType))
						colType = "Nchar";
					else if ("JSON".equals(colType))
						colType = "nvarchar";
					else if ("VARGRAPHIC".equals(colType))
						colType = "nvarchar";
					else if ("timestamp".equals(colType))
						colType = "Datetime2";

					// 供後續Excel使用
					map.put("ColName", colName);
					map.put("ColType", colType);
					map.put("ColLen", colLen);
					map.put("ColNull", colNull);
					map.put("ColPK", colPK);
					mapList.add(map);

					// 組sql用
					colLen = StringUtils.isBlank(colLen) ? "" : "(" + colLen + ")";
					colNull = "Y".equals(colNull) ? "" : "NOT NULL";
					String colSchema = "\n\t" + colName + " " + colType + colLen + " " + colNull + " ,";
					createExternalTable += colSchema;
					createMSSQLTable += colSchema;
				}
				
				createExternalTable = createExternalTable.substring(0,createExternalTable.length()-1) + "\r\n)\r\n"
						+ "WITH (DATA_SOURCE = " + DATA_SOURCE + ", LOCATION = '" + tdDBTableNameDot
						+ "');\r\n\r\n"	;
				createMSSQLTable = createMSSQLTable.substring(0,createMSSQLTable.length()) 
//						+ "\n\tPRIMARY KEY (" + pk + ")\r\n" 
						+ "\n\tCONSTRAINT u_"+tdDBTableNameUL+"_Id UNIQUE (" + pk + ")\r\n" 
						+ ");\r\n\r\n" ;
				
				/**
				 * 與整理過的Table Layout文件比對是否一致
				 */
				runChkTableLayout(mapPath, tdDBTableNameDot, mapTableInfo.get("SubSys"), mapList);
			}

			// 寫出資料移轉所需Script
			FileTools.createFile(mapPath.get("targetSQLPath"), "CreateExternalTable", "sql", createExternalTable);
			FileTools.createFile(mapPath.get("targetSQLPath"), "CreateMSSQLTable", "sql", createMSSQLTable);
			FileTools.createFile(mapPath.get("targetSQLPath"), "InsertInto", "sql", insertInto);
			
		} catch (Exception ex) {
			throw new Exception(className + " Error: \n" + ex);
		}

		System.out.println("runParserSourceSQL 1.解悉出要移轉的清單的Table Layout 2.寫出移轉資料所需的Script Done!");
	}

	/**
	 * 比對Schema與SVN上Table Spec裡的Schema是否一致
	 * 並將Teradata SQL 內的Schema另寫成新的Excel檔，並標註比對結果(紅底表示不一致)
	 * @param mapPath
	 * @param tableName
	 * @param subSys
	 * @param mapListSQLLayout
	 * @throws Exception
	 */
	private static void runChkTableLayout(Map<String,String> mapPath, String tableName, String subSys,
			List<Map<String, String>> mapListSQLLayout) throws Exception {

		Row row;
		Cell cell;

		String sourceTableLayoutExcelPath = mapPath.get("sourceTableLayoutExcelPath");
		String targetTableLayoutExcelPath = mapPath.get("targetTableLayoutExcelPath");
//		boolean isError = false;
		
		try {
			// 找出此檔案放置的確切位置
			String[] folderNameList = new File(sourceTableLayoutExcelPath).list();
			String folderName = "";
			for (String str : folderNameList) {
				if (str.substring(0, str.indexOf("-")).trim().equals(subSys)) {
					folderName = str;
					break;
				}
			}
			sourceTableLayoutExcelPath += folderName + "/";

			String[] fileNameList = new File(sourceTableLayoutExcelPath).list();
			String fileName = "";
			for (String str : fileNameList) {
				if (str.indexOf("-") > 0 && str.substring(0, str.indexOf("-")).trim()
						.equals(tableName.substring(tableName.indexOf(".") + 1).trim())) {
					fileName = str;
					break;
				}
			}
			
			if(StringUtils.isBlank(fileName)) {
				System.out.println("runChkTableLayout " + tableName + " *******Excel 檔案不存在!");
				return;
			}
			sourceTableLayoutExcelPath += fileName;
//			System.out.println("targetTableLayoutExcelPath: "+sourceTableLayoutExcelPath);
			
//			if("C:/22/DW/dw2209/DOCUMENT/3-SD/DW/Table Spec/PMM-責任績效管理/THPMSCR1-績效評分成績檔(預算累計至本月)歷史檔.xlsx".equals(sourceTableLayoutExcelPath))
//				System.out.println("Stop!");
			Sheet sourceSheet = Tools.getSheet(sourceTableLayoutExcelPath, "Layout");
			
			/**
			 * 開始比對Table Layout內容，並將Teradata SQL 內的Schema另寫成新的Excel檔，並標註比對結果(紅底表示不一致)
			 */
			// 因output時需workbook所以多此行只為取workbook
			Workbook targetTableWorkbook = Tools.getWorkbook(targetTableLayoutExcelPath+"../../Sample - Table Layout.xlsx");
			Sheet targetSheet = targetTableWorkbook.getSheet("Layout");
			CellStyle cellStyleNormal = Tools.setStyleNormal(targetTableWorkbook);
			CellStyle cellStyleError = Tools.setStyleError(targetTableWorkbook);
			
			Tools.setCell(tableName.equals(getCellValue(sourceSheet.getRow(0), 4, "TABLE名稱")) ? cellStyleNormal
					: cellStyleError, targetSheet.getRow(0), 4, tableName);
			
			boolean excelEqualsSql = false;
			int lastRowNum = 0;
			for (int i = 4; i <= sourceSheet.getLastRowNum(); i++) {
				row = sourceSheet.getRow(i);
				cell = row == null ? null : row.getCell(0);
				
				if (cell == null || StringUtils.isBlank(cell.toString())) {
					break;
				} else {
					String excelColName = getCellValue(row,1,"欄位名稱").toUpperCase();
					String excelColType = getCellValue(row,3,"資料型態").toUpperCase();
					String excelColLen = getCellValue(row,4,"資料長度").toUpperCase().replace("(", "").replace(")", "").replace(" ", "");
					String excelColNull = getCellValue(row,6,"NULL註記").toUpperCase();
					String excelColPK = getCellValue(row,5,"主鍵註記").toUpperCase();
					
					row = targetSheet.createRow(i);
					excelEqualsSql = false;
					for(Map<String, String> mapLayout : mapListSQLLayout) {
						String sqlColName = mapLayout.get("ColName").toUpperCase();
						if(excelColName.equals(sqlColName)) {
							excelEqualsSql = true;
							String sqlColType = mapLayout.get("ColType").toUpperCase();
							String sqlColLen = mapLayout.get("ColLen").toUpperCase();
							String sqlColNull = mapLayout.get("ColNull").toUpperCase();
							String sqlColPK = mapLayout.get("ColPK").toUpperCase();

							CellStyle sqlColTypeCellStyle = excelColType.equals(sqlColType) ? cellStyleNormal : cellStyleError;
							CellStyle sqlColLenCellStyle = excelColLen.equals(sqlColLen) ? cellStyleNormal : cellStyleError;
							CellStyle sqlColNullCellStyle = excelColNull.equals(sqlColNull) ? cellStyleNormal : cellStyleError;
							CellStyle sqlColPKCellStyle = excelColPK.equals(sqlColPK) ? cellStyleNormal	: cellStyleError;
							
//							isError = (sqlColTypeCellStyle.equals(cellStyleError)
//									|| sqlColLenCellStyle.equals(cellStyleError)
//									|| sqlColNullCellStyle.equals(cellStyleError)
//									|| sqlColPKCellStyle.equals(cellStyleError)) ? true : false;
							
							cell = row.createCell(0);
							cell.setCellFormula("ROW()-4");
							cell.setCellStyle(cellStyleNormal);
							Tools.setCell(cellStyleNormal, row, 1, sqlColName);
							Tools.setCell(sqlColTypeCellStyle,  row, 3, sqlColType);
							Tools.setCell(sqlColLenCellStyle,  row, 4, sqlColLen);
							Tools.setCell(sqlColNullCellStyle,  row, 6, sqlColNull);
							Tools.setCell(sqlColPKCellStyle,  row, 5, sqlColPK);
							break;
						}
					}
					// 若Excel內的欄位名稱比對不到SQL的欄位則另執行此段
					if(!excelEqualsSql) {
						cell = row.createCell(0);
						cell.setCellFormula("ROW()-4");
						cell.setCellStyle(cellStyleError);
						Tools.setCell(cellStyleError, row, 1, excelColName);
						Tools.setCell(cellStyleError, row, 2, "(Script無此欄位)");
						Tools.setCell(cellStyleError, row, 3, "");
						Tools.setCell(cellStyleError, row, 4, "");
						Tools.setCell(cellStyleError, row, 5, "");
						Tools.setCell(cellStyleError, row, 6, "");
//						isError = true;
					}
					lastRowNum = i;
				}
			}
			
			// 以SQL內的欄位名稱為主去比對Excel，找出是否SQL有但Excel沒有的欄位
			lastRowNum++;
			for(Map<String, String> mapLayout : mapListSQLLayout) {
				excelEqualsSql = false;
				String sqlColName = mapLayout.get("ColName").toUpperCase();
				for (int i = 4; i <= sourceSheet.getLastRowNum(); i++) {
					row = sourceSheet.getRow(i);
					cell = row == null ? null : row.getCell(0);
					if (cell == null || StringUtils.isBlank(cell.toString())) {
						break;
					} else {
						String excelColName = getCellValue(row,1,"欄位名稱").toUpperCase();
						if(excelColName.equals(sqlColName)) {
							excelEqualsSql = true;
							break;
						}
					}
				}
				if(!excelEqualsSql) {
					row = targetSheet.createRow(lastRowNum++);
					cell = row.createCell(0);
					cell.setCellFormula("ROW()-4");
					cell.setCellStyle(cellStyleError);
					Tools.setCell(cellStyleError, row, 1, sqlColName);
					Tools.setCell(cellStyleError, row, 2, "");
					Tools.setCell(cellStyleError, row, 3, mapLayout.get("ColType").toUpperCase());
					Tools.setCell(cellStyleError, row, 4, mapLayout.get("ColLen").toUpperCase());
					Tools.setCell(cellStyleError, row, 5, mapLayout.get("ColNull").toUpperCase());
					Tools.setCell(cellStyleError, row, 6, mapLayout.get("ColNull").toUpperCase());
//					isError = true;
				}
			}
			
			// 將整理好的比對結果另寫出Excel檔
			Tools.output(targetTableWorkbook, "2007", targetTableLayoutExcelPath, "Target - " + subSys + " " + fileName);

//			System.out.println("runChkTableLayout " + fileName + " 檔案比對結果: " + (isError ? "不一致" : "ok"));
		} catch (Exception ex) {
			throw new Exception(className + "runChkTableLayout Error: \n" + ex);
		}
		
	}

	/**
	 * 將連兩個空格取代為一個空格
	 * @param str
	 * @return
	 * @throws Exception
	 */
	private static String replaceDoubleSpace(String str) throws Exception {
		for(int i = 0;i< 10;i++) {
			str = str.replace("  ", " ");
		}
		return str;
	}
	
	/**
	 * 取Excel欄位值
	 * 
	 * @param sheet
	 * @param rownum
	 * @param cellnum
	 * @param fieldName
	 * @return
	 */
	private static String getCellValue(Row row, int cellnum, String fieldName) throws Exception {
		try {
			Cell cell = row.getCell(cellnum);
			if (!cellNotBlank(cell))
				return "";
			else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
				return String.valueOf((int) cell.getNumericCellValue()).trim();
			else if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue().trim();
			else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				if (cell.getCachedFormulaResultType() == Cell.CELL_TYPE_NUMERIC)
					return String.valueOf((int) cell.getNumericCellValue()).trim();
				else if (cell.getCellType() == Cell.CELL_TYPE_STRING)
					return cell.getStringCellValue().trim();
			}
		} catch (Exception ex) {
			throw new Exception(className + " getCellValue " + fieldName + " 格式錯誤");
		}
		return "";
	}
	
	/**
     * 不為空
     */
	private static boolean cellNotBlank(Cell cell) {
		return cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK;
	}
}
