package com.jin.commons.poi;

import com.jin.commons.poi.model.SheetSettings;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: OfficeIoUtils
 * @Description: OFFICE导入导出工具类
 * @author wujinglei
 * @date 2014年6月10日 下午5:57:20
 *
 */
public final class OfficeIoUtils {

	private final static Logger log = LoggerFactory.getLogger(OfficeIoFactory.class);

	private final static OfficeIoFactory IO_FACTORY = new OfficeIoFactory();

	/**
	 * 导出XLS模板
	 * @param sheets
	 * @return
	 */
	public static OfficeIoResult exportXlsxTemplate(SheetSettings sheets){
		return IO_FACTORY.exportXlsxTemplate(sheets);
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午11:38:40
	 * @Description: 导出模板
	 * @param sheets
	 * @return
	 */
	public static OfficeIoResult exportXlsxTemplate(SheetSettings[] sheets){
		return IO_FACTORY.exportXlsxTemplate(sheets);
	}

	/**
	 * 导出Xlsx
	 * @param sheetSettings
	 * @return
	 */
	public static OfficeIoResult exportXlsx(SheetSettings sheetSettings){
		return IO_FACTORY.exportXlsx(new SheetSettings[]{sheetSettings});
	}

	/**
	 * 导出Xlsx
	 * @param sheetSettingsArray
	 * @return
	 */
	public static OfficeIoResult exportXlsx(SheetSettings[] sheetSettingsArray){
		return IO_FACTORY.exportXlsx(sheetSettingsArray);
	}

	/**
	 * 导入Xlsx
	 * @param inputStream
	 * @param sheets
	 * @return
	 */
	public static OfficeIoResult importXlsx(InputStream inputStream, SheetSettings[] sheets) {
		return IO_FACTORY.importXlsx(inputStream, sheets);
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月11日 上午10:29:30
	 * @Description: 导入
	 * @param file
	 * @param sheets
	 * @return
	 */
	public static OfficeIoResult importXlsx(File file, SheetSettings[] sheets) {
		return IO_FACTORY.importXlsx(file, sheets);
	}

	/**
	 * 导入Xlsx
	 * @param inputStream
	 * @param sheetSettings
	 * @return
	 */
	public static OfficeIoResult importXlsx(InputStream inputStream, SheetSettings sheetSettings) {
		return IO_FACTORY.importXlsx(inputStream, new SheetSettings[]{sheetSettings});
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月11日 上午10:29:30
	 * @Description: 导入
	 * @param file
	 * @param sheetSettings
	 * @return
	 */
	public static OfficeIoResult importXlsx(File file, SheetSettings sheetSettings) {
		return IO_FACTORY.importXlsx(file, new SheetSettings[]{sheetSettings});
	}

	/**
	 * @author: wujinglei
	 * @date: 2014-6-20 下午3:50:03
	 * @Description: 导出出错信息内容
	 * @param sheets
	 * @param errRecordRows
	 * @return
	 */
	public static OfficeIoResult exportErrorRecord(SheetSettings[] sheets, Map<Integer,List> errRecordRows){
		return IO_FACTORY.exportXlsxErrorRecord(sheets, errRecordRows);
	}

	/**
	 * @author: wujinglei
	 * @date: 2014-6-21 下午1:14:44
	 * @Description: 导出异常数据
	 * @param sheets
	 * @param errRecordRows
	 * @param filePath
	 * @return
	 */
	public static boolean exportErrorFile(SheetSettings[] sheets,Map<Integer,List> errRecordRows, String filePath){
		OfficeIoResult errResult = IO_FACTORY.exportXlsxErrorRecord(sheets, errRecordRows);
		FileOutputStream output = null;
		try {
			output = new FileOutputStream(new File(filePath));
			errResult.getResultWorkbook().write(output);
			return true;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return false;
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		} finally{
			try {
				output.close();
			} catch (IOException e) {
				e.printStackTrace();
			}  	    
		}
	}
}
