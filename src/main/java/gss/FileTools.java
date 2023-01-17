package gss;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;

public class FileTools {

	// 檔案路徑 名稱
	private static String filenameTemp;

	/**
	 * 建立檔案
	 * 
	 * @param path			檔路徑
	 * @param fileName		檔名稱
	 * @param extension		副檔名
	 * @param fileContent	檔案內容
	 * @return 是否建立成功，成功則返回true
	 */
	public static boolean createFile(String path, String fileName, String extension, String fileContent) {
		Boolean bool = false;
		File file ;
		
		try {
			file = new File(path);
			if(!file.exists()) file.mkdirs();
			
			filenameTemp = path + fileName + "." + extension;// 檔案路徑 名稱 檔案型別
			file = new File(filenameTemp);
			// 如果檔案不存在，則建立新的檔案
			if (!file.exists()) {
				file.createNewFile();
				bool = true;
				System.out.println("success create file: " + filenameTemp);
			}
			// 建立檔案成功後，寫入內容到檔案裡
			writeFileContent(filenameTemp, fileContent);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return bool;
	}

	/**
	 * 向檔案中寫入內容
	 * 
	 * @param filePathName 檔名稱
	 * @param newstr   寫入的內容
	 * @return
	 * @throws IOException
	 */
	public static boolean writeFileContent(String filePathName, String newstr) throws IOException {
		Boolean bool = false;
//		String filein = "\r\n" + newstr + "\r\n";// 新寫入的行，換行
//		String temp = "";
//		FileInputStream fis = null;
//		InputStreamReader isr = null;
//		BufferedReader br = null;
		FileOutputStream fos = null;
		PrintWriter pw = null;
		try {
			File file = new File(filePathName);// 檔案路徑(包括檔名稱)
//			// 將原檔案內容讀入輸入流
//			fis = new FileInputStream(file);
//			isr = new InputStreamReader(fis);
//			br = new BufferedReader(isr);
			StringBuffer buffer = new StringBuffer();
//			// 寫入檔案原有內容
//			while((temp = br.readLine()) != null) {
//				buffer.append(temp);
//				// 行與行之間的分隔符 相當於“\n”
//				buffer = buffer.append(System.getProperty("line.separator"));
//			}
			buffer.append(newstr);
			fos = new FileOutputStream(file);
			pw = new PrintWriter(fos);
			pw.write(buffer.toString().toCharArray());
			pw.flush();
			bool = true;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (pw != null)	pw.close();
			if (fos != null) fos.close();
//			if (br != null) br.close();
//			if (isr != null) isr.close();
//			if (fis != null) fis.close();
		}
		return bool;
	}

	/**
	 * 讀取檔案內容
	 */
	public static String readFileContent(String filePathName) throws IOException {
		String temp = "";
		FileInputStream fis = null;
		InputStreamReader isr = null;
		BufferedReader br = null;
		
		try {
			File file = new File(filePathName);// 檔案路徑(包括檔名稱)
			// 將檔案內容讀入輸入流
			fis = new FileInputStream(file);
			isr = new InputStreamReader(fis);
			br = new BufferedReader(isr);
			StringBuffer buffer = new StringBuffer();
			// 讀取檔案內容
			while((temp = br.readLine()) != null) {
				buffer.append(temp);
				// 行與行之間的分隔符 相當於“\n”
				buffer.append(System.getProperty("line.separator"));
			}
			
//			System.out.println(buffer);
			return buffer.toString();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (br != null) br.close();
			if (isr != null) isr.close();
			if (fis != null) fis.close();
		}
		return "";
	}
	
	/**
	 * 刪除檔案
	 * 
	 * @param fileName 檔名稱
	 * @return
	 */
	public static boolean delFile(String path, String fileName) {
		Boolean bool = false;
		filenameTemp = path + fileName + ".txt";
		File file = new File(filenameTemp);
		try {
			if (file.exists()) {
				file.delete();
				bool = true;
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
		return bool;
	}

}
