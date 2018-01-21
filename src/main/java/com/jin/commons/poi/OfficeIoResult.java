package com.jin.commons.poi;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.jin.commons.poi.model.WrongRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.jin.commons.poi.model.ErrorRecord;
import com.jin.commons.poi.model.SheetSettings;

/**
 * The type Office io result.
 *
 * @param <T> the type parameter
 * @author wujinglei
 * @ClassName: OfficeIoResult
 * @Description: office导入导出结果
 * @date 2014年6月11日 上午9:53:03
 */
public final class OfficeIoResult<T> {
	
	/**
	 * 出错记录信息
	 */
	private final List<ErrorRecord> errors = new ArrayList<ErrorRecord>();

    /**
     * 警告记录信息
     */
    private final List<WrongRecord> wrongs = new ArrayList<WrongRecord>();
	
	/**
	 * 出错行对象
	 */
	private final Map<Integer,List> errRecordRows = new HashMap<Integer,List>();
	
	/**
	 * 导入生成的结果集
	 */
	private final List<T> importList = new ArrayList<T>();
	
	/**
	 * 原始数据
	 */
	private final List originalList = new ArrayList();
	
	/**
	 * 导出结果集
	 */
	private final XSSFWorkbook resultWorkbook = new XSSFWorkbook();

	private SheetSettings[] sheetSettings;
	
	/**
	 * 返回成功的结果条数
	 */
	private Long[] resultTotal;
	
	/**
	 * 文件中的行数
	 */
	private Long[] fileTotalRow;
	
	/**
	 * msg记录导入数据的消息
	 */
	private String msg;

	private Boolean isCompleted = true;

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 下午1:10:45
	 * @Description:
	 */
	private OfficeIoResult(){
		
	}

	/**
	 * Instantiates a new Office io result.
	 *
	 * @param sheets the sheets
	 * @author: wujinglei
	 * @date: 2014年6月12日 下午1:10:54
	 * @Description:
	 */
	public OfficeIoResult(SheetSettings[] sheets ){
		if (sheets != null){
			resultTotal = new Long[sheets.length];
		}
		if(sheets != null && sheets.length > 0){
			for (SheetSettings sheetItem : sheets) {
				originalList.add(sheetItem.getExportData());
			}
		}
	}

	/**
	 * Get original list list.
	 *
	 * @return the list
	 */
	public final List getOriginalList(){
		return originalList;
	}

	/**
	 * Get result total long [ ].
	 *
	 * @return the resultTotal
	 */
	public Long[] getResultTotal() {
		return resultTotal;
	}

	/**
	 * Sets result total.
	 *
	 * @param resultTotal the resultTotal to set
	 */
	public void setResultTotal(Long[] resultTotal) {
		this.resultTotal = resultTotal;
	}

	/**
	 * Gets errors.
	 *
	 * @return the errors
	 */
	public List<ErrorRecord> getErrors() {
		return errors;
	}

	/**
	 * Gets wrongs.
	 *
	 * @return the wrongs
	 */
	public List<WrongRecord> getWrongs() {
        return wrongs;
    }

	/**
	 * Gets import list.
	 *
	 * @return the importList
	 */
	public List<T> getImportList() {
		return importList;
	}

	/**
	 * Add sheet list.
	 *
	 * @param sheetList the sheet list
	 */
	public void addSheetList(List<T> sheetList){
		importList.addAll(sheetList);
	}

	/**
	 * Gets result workbook.
	 *
	 * @return the resultWorkbook
	 */
	public XSSFWorkbook getResultWorkbook() {
		return resultWorkbook;
	}

	/**
	 * Add error record.
	 *
	 * @param errorRecord the error record
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:36:24
	 * @Description: 将ErrorRecord添加 到列表中
	 */
	public void addErrorRecord(ErrorRecord errorRecord){
        this.errors.add(errorRecord);
    }

	/**
	 * Add wrong record.
	 *
	 * @param wrongRecord the wrong record
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:36:24
	 * @Description: 将WrongRecord添加 到列表中
	 */
	public void addWrongRecord(WrongRecord wrongRecord){
        this.wrongs.add(wrongRecord);
    }

	/**
	 * Gets err record rows.
	 *
	 * @return the err record rows
	 */
	public Map<Integer, List> getErrRecordRows() {
		return errRecordRows;
	}

	/**
	 * Add error record row.
	 *
	 * @param index    the index
	 * @param errorRow the error row
	 * @author: wujinglei
	 * @date: 2014 -6-20 下午2:00:44
	 * @Description: 将行记录放入errorRecordRow中
	 */
	public void addErrorRecordRow(Integer index,Row errorRow){
		List targetList = this.errRecordRows.get(index);
		if (targetList == null){
			targetList = new ArrayList<Row>();
			this.errRecordRows.put(index,targetList);
		}
		targetList.add(errorRow);
	}

	/**
	 * Add error record row.
	 *
	 * @param index the index
	 * @param strs  the strs
	 * @author: wujinglei
	 * @date: 2014 -6-20 下午2:00:44
	 * @Description: 将行记录放入errorRecordRow中
	 */
	public void addErrorRecordRow(Integer index,String[] strs){
		List targetList = this.errRecordRows.get(index);
		if (targetList == null){
			targetList = new ArrayList<Row>();
			this.errRecordRows.put(index,targetList);
		}
		targetList.add(strs);
	}

	/**
	 * Print error record string.
	 *
	 * @return string
	 * @author: wujinglei
	 * @date: 2014 -6-25 上午11:30:12
	 * @Description: 打印异常记录
	 */
	public String printErrorRecord(){
		StringBuffer errStr = new StringBuffer("");
		for (ErrorRecord errorRecord : this.errors) {
			errStr.append(errorRecord.toString() + "\n<br>");
		}
		return errStr.toString();
	}

	/**
	 * Print wrong record string.
	 *
	 * @return string
	 * @author: wujinglei
	 * @date: 2014 -6-25 上午11:30:12
	 * @Description: 打印警告记录
	 */
	public String printWrongRecord(){
        StringBuffer wrongStr = new StringBuffer("");
        for (WrongRecord wrongRecord : this.wrongs) {
            wrongStr.append(wrongRecord.toString() + "\n<br>");
        }
        return wrongStr.toString();
    }

	/**
	 * Get file total row long [ ].
	 *
	 * @return the fileTotalRow
	 */
	public Long[] getFileTotalRow() {
		return fileTotalRow;
	}

	/**
	 * Sets file total row.
	 *
	 * @param fileTotalRow the fileTotalRow to set
	 */
	public void setFileTotalRow(Long[] fileTotalRow) {
		this.fileTotalRow = fileTotalRow;
	}

	/**
	 * Set total row count.
	 *
	 * @param sheetIndex the sheet index
	 * @param rows       the rows
	 * @author: wujinglei
	 * @date: 2014 -6-25 上午11:30:30
	 * @Description: 将总数量进行赋值
	 */
	public void setTotalRowCount(int sheetIndex,Long rows){
		this.fileTotalRow[sheetIndex] = rows;
	}

	/**
	 * Set result count.
	 *
	 * @param sheetIndex the sheet index
	 * @param rows       the rows
	 * @author: wujinglei
	 * @date: 2014 -6-25 上午11:30:54
	 * @Description: 将最终总数量进行赋值
	 */
	public void setResultCount(int sheetIndex,Long rows){
		this.resultTotal[sheetIndex] = rows;
	}

	/**
	 * Gets msg.
	 *
	 * @return the msg
	 */
	public String getMsg() {
		return msg;
	}

	/**
	 * Sets msg.
	 *
	 * @param msg the msg
	 */
	public void setMsg(String msg) {
		this.msg = msg;
	}

	public Boolean isCompleted() {
		return isCompleted;
	}

	public void setCompleted(Boolean completed) {
		isCompleted = completed;
	}

	public SheetSettings[] getSheetSettings() {
		return sheetSettings;
	}

	void setSheetSettings(SheetSettings[] sheetSettings) {
		this.sheetSettings = sheetSettings;
	}
}
