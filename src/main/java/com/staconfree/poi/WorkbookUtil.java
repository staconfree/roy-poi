package com.staconfree.poi;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 
 * @author xiexq
 * @date 2014年4月29日
 * @since 1.0
 */
public class WorkbookUtil {

	public static void flush(Workbook workbook, String writeFilePath) throws IOException {
		// 创建目录
		File descFile = new File(writeFilePath);
		if (!descFile.getParentFile().exists()) {
			descFile.getParentFile().mkdirs();
		}
		FileOutputStream out = new FileOutputStream(writeFilePath);
		workbook.write(out);
		if (out != null) {
			try {
				out.close();
				out = null;
			} catch (IOException e) {
				out = null;
				e.printStackTrace();
			}
		}
	}

}
