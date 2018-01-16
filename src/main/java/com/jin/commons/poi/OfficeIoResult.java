package com.jin.commons.poi;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.jin.commons.poi.model.ErrorRecord;
import com.jin.commons.poi.model.SheetOptions;

/**
 * @ClassName: OfficeIoResult
 * @Description: office导入导出结果
 * @author wujinglei
 * @date 2014年6月11日 上午9:53:03
 *
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
	private final List<T> importList = new ArrayList();
	
	/**
	 * 原始数据
	 */
	private final List originalList = new ArrayList();
	
	/**
	 * 导出结果集
	 */
	private final XSSFWorkbook resultWorkbook = new XSSFWorkbook();
	
	/**
	 * 导出97-03版本
	 */
	private final HSSFWorkbook wb = new HSSFWorkbook();

	/**
	 * 返回成功的结果条数
	 */
	private Long[] resultTotal;
	
	/**
	 * 文件中的行数
	 */
	private Long[] fileTotalRow;
	
	/**
	 * KEYID记录导入数据最关键的ID
	 */
	private String keyId;
	
	/**
	 * msg记录导入数据的消息
	 */
	private String msg;

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 下午1:10:45
	 * @Description:
	 */
	@SuppressWarnings("unused")
	private OfficeIoResult(){
		
	}
	
	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 下午1:10:54
	 * @Description:
	 */
	public OfficeIoResult(SheetOptions[] sheets ){
		if (sheets != null){
			resultTotal = new Long[sheets.length];
		}
		if(sheets != null && sheets.length > 0){
			for (SheetOptions sheetItem : sheets) {
				originalList.add(sheetItem.getExportData());
			}
		}
	}
	
	public final List getOriginalList(){
		return originalList;
	}
	
	/**
	 * @return the resultTotal
	 */
	public Long[] getResultTotal() {
		return resultTotal;
	}

	/**
	 * @param resultTotal the resultTotal to set
	 */
	public void setResultTotal(Long[] resultTotal) {
		this.resultTotal = resultTotal;
	}

	/**
	 * @return the errors
	 */
	public List<ErrorRecord> getErrors() {
		return errors;
	}

    /**
     * @return the wrongs
     */
    public List<WrongRecord> getWrongs() {
        return wrongs;
    }

	/**
	 * @return the importList
	 */
	public List<T> getImportList() {
		return importList;
	}
	
	public void addSheetList(List sheetList){
		importList.addAll(sheetList);
	}
	/**
	 * @return the resultWorkbook
	 */
	public XSSFWorkbook getResultWorkbook() {
		return resultWorkbook;
	}

    /**
     * @author: wujinglei
     * @date: 2014年6月12日 上午10:36:24
     * @Description: 将ErrorRecord添加 到列表中
     * @param errorRecord
     */
    public void addErrorRecord(ErrorRecord errorRecord){
        this.errors.add(errorRecord);
    }

    /**
     * @author: wujinglei
     * @date: 2014年6月12日 上午10:36:24
     * @Description: 将WrongRecord添加 到列表中
     * @param wrongRecord
     */
    public void addWrongRecord(WrongRecord wrongRecord){
        this.wrongs.add(wrongRecord);
    }
	
	public Map<Integer, List> getErrRecordRows() {
		return errRecordRows;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014-6-20 下午2:00:44
	 * @Description: 将行记录放入errorRecordRow中
	 * @param index
	 * @param errorRow
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
	 * @author: wujinglei
	 * @date: 2014-6-20 下午2:00:44
	 * @Description: 将行记录放入errorRecordRow中
	 * @param index
	 * @param strs
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
	 * @author: wujinglei
	 * @date: 2014-6-25 上午11:30:12
	 * @Description: 打印异常记录
	 * @return
	 */
	public String printErrorRecord(){
		StringBuffer errStr = new StringBuffer("");
		for (ErrorRecord errorRecord : this.errors) {
			errStr.append(errorRecord.toString() + "\n<br>");
		}
		return errStr.toString();
	}

    /**
     * @author: wujinglei
     * @date: 2014-6-25 上午11:30:12
     * @Description: 打印警告记录
     * @return
     */
    public String printWrongRecord(){
        StringBuffer wrongStr = new StringBuffer("");
        for (WrongRecord wrongRecord : this.wrongs) {
            wrongStr.append(wrongRecord.toString() + "\n<br>");
        }
        return wrongStr.toString();
    }

	/**
	 * @return the fileTotalRow
	 */
	public Long[] getFileTotalRow() {
		return fileTotalRow;
	}

	/**
	 * @param fileTotalRow the fileTotalRow to set
	 */
	public void setFileTotalRow(Long[] fileTotalRow) {
		this.fileTotalRow = fileTotalRow;
	}
	
	/**
	 * @author: wujinglei
	 * @date: 2014-6-25 上午11:30:30
	 * @Description: 将总数量进行赋值
	 * @param sheetIndex
	 * @param rows
	 */
	public void setTotalRowCount(int sheetIndex,Long rows){
		this.fileTotalRow[sheetIndex] = rows;
	}
	
	/**
	 * @author: wujinglei
	 * @date: 2014-6-25 上午11:30:54
	 * @Description: 将最终总数量进行赋值
	 * @param sheetIndex
	 * @param rows
	 */
	public void setResultCount(int sheetIndex,Long rows){
		this.resultTotal[sheetIndex] = rows;
	}

	/**
	 * @return the keyId
	 */
	public String getKeyId() {
		return keyId;
	}

	/**
	 * @param keyId the keyId to set
	 */
	public void setKeyId(String keyId) {
		this.keyId = keyId;
	}

	public String getMsg() {
		return msg;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

	public HSSFWorkbook getWb() {
		return wb;
	}



}
