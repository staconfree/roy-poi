package com.staconfree.poi.write.sxssf;

import com.staconfree.clazz.model.Student;
import com.staconfree.clazz.model.Teacher;
import com.staconfree.poi.write.DataLoopModelHelper;
import com.staconfree.poi.write.ExcelParser;
import com.staconfree.poi.write.WriteCallBack;
import com.staconfree.poi.write.WriteCellHelper;
import com.staconfree.poi.write.common.SimpleCopyWorkBook;
import com.staconfree.poi.write.common.SimpleDataLoopModelHelper;
import com.staconfree.poi.write.common.SimpleInsertFooter;
import com.staconfree.poi.write.common.SimpleInsertLoopData;
import com.staconfree.poi.write.model.DataLoopModel;
import com.staconfree.poi.write.model.FooterModel;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月7日
 * @since 1.0
 */
public class SxssfExcelParser implements ExcelParser {
	private final Log log = LogFactory.getLog(getClass());
	/**
	 * 模板路径
	 */
	private String templateFilePath;
	/**
	 * 模板sheetName
	 */
	private String templateSheetName;
	/**
	 * 写文件路径
	 */
	private String writeFilePath;

	private SXSSFWorkbook workbook = new SXSSFWorkbook(50);
	private XSSFWorkbook srcWorkbook;
	
	private WriteCellHelper writeCellHelper = new SxssfWriteCellHelper();
	
	private WriteCallBack writeCallBack;
	/**
	 * 循环区域的配置
	 */
	private List<DataLoopModel> dataLoopModels = new ArrayList<DataLoopModel>();

	/**
	 * 循环区域下面的内容
	 */
	private FooterModel footerModel = new FooterModel();
	/**
	 * 循环体的合并区域
	 */
	private List<CellRangeAddress> loopRangeAddresses = new ArrayList<CellRangeAddress>();

	public SxssfExcelParser() {

	}

	public SxssfExcelParser(String templateFilePath, String writeFilePath) {
		setTemplateFilePath(templateFilePath);
		setWriteFilePath(writeFilePath);
	}

	@Override
	public void parse(Map<String, Object> dataStack) throws Exception {
		if (templateFilePath == null || templateSheetName == null || writeFilePath == null) {
			return;
		}
		// 创建目录
		File descFile = new File(writeFilePath);
		if (!descFile.getParentFile().exists()) {
			descFile.getParentFile().mkdirs();
		}
		InputStream ins = null;
		FileOutputStream out = null;
		InputStream newIns = null;
		try {
			ins = new FileInputStream(new File(templateFilePath));
			if (workbook.getSheet(templateSheetName) == null) {// 拷贝模板，并赋值一部分数据
				srcWorkbook = new XSSFWorkbook(ins);
				dataLoopModels = new ArrayList<DataLoopModel>();
				footerModel = new FooterModel();
				loopRangeAddresses = new ArrayList<CellRangeAddress>();
				SimpleCopyWorkBook xssfCopyWorkBookHelper = new SimpleCopyWorkBook();
				xssfCopyWorkBookHelper.setWriteCellHelper(writeCellHelper);
				xssfCopyWorkBookHelper.setWriteCallBack(writeCallBack);
				xssfCopyWorkBookHelper.copy(srcWorkbook, templateSheetName, workbook, templateSheetName,
						dataLoopModels, footerModel, dataStack);
			}
			// 插入循环数据
			insertLoopData(dataStack);
		} catch (Exception e) {
			log.error("", e);
			throw new Exception(e);
		} finally {
			if (ins != null) {
				try {
					ins.close();
					ins = null;
				} catch (IOException e) {
					ins = null;
					e.printStackTrace();
				}
			}
			if (out != null) {
				try {
					out.close();
					out = null;
				} catch (IOException e) {
					out = null;
					e.printStackTrace();
				}
			}
			if (newIns != null) {
				try {
					newIns.close();
					newIns = null;
				} catch (IOException e) {
					newIns = null;
					e.printStackTrace();
				}
			}
		}

	}

	private void insertLoopData(Map<String, Object> dataStack) throws Exception {
		try {
			SimpleInsertLoopData insertLoopData = new SimpleInsertLoopData();
			insertLoopData.setWriteCallBack(writeCallBack);
			insertLoopData.setWriteCellHelper(writeCellHelper);
			if (dataLoopModels != null && !dataLoopModels.isEmpty()) {
				DataLoopModelHelper dataLoopModelHelper = new SimpleDataLoopModelHelper();
				dataLoopModelHelper.setWriteCellHelper(writeCellHelper);
				for (DataLoopModel dataLoopModel : dataLoopModels) {
					for (String stackKey : dataStack.keySet()) {
						if (stackKey.equals(dataLoopModel.getName())) {
							Sheet sheet = workbook.getSheet(templateSheetName);
							// 从dataLoopModel的bgnRow插入数据
							int insertNum = insertLoopData.insert(sheet, (List) dataStack.get(stackKey), dataLoopModel,
									dataStack, loopRangeAddresses);
							// 更新所有dataLoopModel的bgnRow
							dataLoopModelHelper.updateBgnRow(dataLoopModels, dataLoopModel.getBgnRow(), insertNum);
						}
					}
				}
			}
		} catch (Exception e) {
			log.error("", e);
			throw new Exception(e);
		}
	}

	private void insertFooter(int insertLoopNum) throws Exception {
		Sheet sheet = workbook.getSheet(templateSheetName);
		SimpleInsertFooter insertFooter = new SimpleInsertFooter();
		insertFooter.setWriteCellHelper(writeCellHelper);
		insertFooter.setWriteCallBack(writeCallBack);
		insertFooter.insert(workbook, sheet, footerModel, insertLoopNum);
	}

	@Override
	public void endSheet() throws Exception {
		if (writeFilePath == null || (footerModel == null && loopRangeAddresses.isEmpty())) {
			return;
		}
		try {
			// 插入文件尾
			int totalInsertNum = 0;
			for (DataLoopModel dataLoopModel : dataLoopModels) {
				totalInsertNum += dataLoopModel.getLoopedNum();
			}
			insertFooter(totalInsertNum);
			// 最后再处理合并区域
			// 1、处理循环体的合并区域
			Sheet sheet = workbook.getSheet(templateSheetName);
			if (sheet != null) {
				for (CellRangeAddress cellRangeAddress : loopRangeAddresses) {
					sheet.addMergedRegion(cellRangeAddress);
				}
				// 2、处理footer的合并区域
				if (footerModel != null) {
					List<CellRangeAddress> footerRangeAddresses = footerModel.getCellRangeAddresses();
					for (CellRangeAddress cellRangeAddress : footerRangeAddresses) {
						sheet.addMergedRegion(cellRangeAddress);
					}
				}
			}
		} catch (Exception e) {
			log.error("", e);
			throw new Exception(e);
		} finally {

		}
	}

	public void flush() throws Exception {
		if (writeFilePath == null || workbook == null) {
			return;
		}
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(writeFilePath);
			workbook.write(out);
			workbook = null;
		} catch (Exception e) {
			e.printStackTrace();
			throw new Exception(e);
		} finally {
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

	@Override
	public void setTemplateFilePath(String templateFilePath) {
		this.templateFilePath = templateFilePath;
	}

	@Override
	public void setTemplateSheetName(String templateSheetName) {
		this.templateSheetName = templateSheetName;
	}

	@Override
	public void setWriteFilePath(String writeFilePath) {
		this.writeFilePath = writeFilePath;
	}

	public WriteCallBack getWriteCallBack() {
		return writeCallBack;
	}

	public void setWriteCallBack(WriteCallBack writeCallBack) {
		this.writeCallBack = writeCallBack;
	}

	public WriteCellHelper getWriteCellHelper() {
		return writeCellHelper;
	}

	public void setWriteCellHelper(WriteCellHelper writeCellHelper) {
		this.writeCellHelper = writeCellHelper;
	}

}
