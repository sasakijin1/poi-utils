package com.jin.commons.poi.model;

import java.io.Serializable;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @ClassName: ErrorRecord
 * @Description: 出错记录
 * @author wujinglei
 * @date 2014年6月11日 上午9:54:54
 *
 */
public class ErrorRecord implements Serializable{

	private static final long serialVersionUID = -1720798939182303079L;

	private final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	
	/**
	 * sheet 名称
	 */
	private String sheetName;
	
	/**
	 * 编码级错误
	 */
	private boolean codeError;
	
	/**
	 * 坐标号
	 */
	private String adress;
	
	/**
	 * 列参数
	 */
	private CellSettings cellSettings;
	
	/**
	 * 出错信息
	 */
	private String errorMsg;
	
	/**
	 * 处理方式
	 */
	private String handleType;
	
	/**
	 * 记录时间
	 */
	private Date recordTime = new Date();
	
	@Override
	public String toString() {
		StringBuffer positionStr = new StringBuffer();
		positionStr.append(sdf.format(recordTime));
		positionStr.append(":表号:").append(sheetName).append(",位置:").append(adress);
		positionStr.append(",异常信息:").append(errorMsg);
		positionStr.append(",处理方式:").append(handleType);
		return positionStr.toString();
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:34:56
	 * @Description:
	 */
	private ErrorRecord(){
		
	}
	
	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:55:48
	 * @Description:
	 */
	public ErrorRecord(String errorMsg, String handleType,boolean codeError) {
		this.errorMsg = errorMsg;
		this.handleType = handleType;
		this.codeError = codeError;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午11:08:42
	 * @Description:
	 */
	public ErrorRecord(String sheetName,String adress, String errorMsg, String handleType,boolean codeError) {
		super();
		this.sheetName = sheetName;
		this.adress = adress;
		this.errorMsg = errorMsg;
		this.handleType = handleType;
		this.codeError = codeError;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:34:11
	 * @Description:
	 */
	public ErrorRecord(String sheetName,String adress, CellSettings cellSettings, String errorMsg, String handleType,boolean codeError) {
		this.sheetName = sheetName;
		this.adress = adress;
		this.cellSettings = cellSettings;
		this.errorMsg = errorMsg;
		this.handleType = handleType;
		this.codeError = codeError;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午11:05:04
	 * @Description:
	 */
	public ErrorRecord(String sheetName, String errorMsg, String handleType,boolean codeError) {
		super();
		this.sheetName = sheetName;
		this.errorMsg = errorMsg;
		this.handleType = handleType;
		this.codeError = codeError;
	}

	public SimpleDateFormat getSdf() {
		return sdf;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public String getAdress() {
		return adress;
	}

	public void setAdress(String adress) {
		this.adress = adress;
	}

	/**
	 * @return the cellSettings
	 */
	public CellSettings getCellSettings() {
		return cellSettings;
	}

	/**
	 * @param cellSettings the cellSettings to set
	 */
	public void setCellSettings(CellSettings cellSettings) {
		this.cellSettings = cellSettings;
	}

	/**
	 * @return the errorMsg
	 */
	public String getErrorMsg() {
		return errorMsg;
	}

	/**
	 * @param errorMsg the errorMsg to set
	 */
	public void setErrorMsg(String errorMsg) {
		this.errorMsg = errorMsg;
	}

	/**
	 * @return the handleType
	 */
	public String getHandleType() {
		return handleType;
	}

	/**
	 * @param handleType the handleType to set
	 */
	public void setHandleType(String handleType) {
		this.handleType = handleType;
	}

	/**
	 * @return the recordTime
	 */
	public Date getRecordTime() {
		return recordTime;
	}

	/**
	 * @param recordTime the recordTime to set
	 */
	public void setRecordTime(Date recordTime) {
		this.recordTime = recordTime;
	}

	/**
	 * @return the codeError
	 */
	public boolean isCodeError() {
		return codeError;
	}

	/**
	 * @param codeError the codeError to set
	 */
	public void setCodeError(boolean codeError) {
		this.codeError = codeError;
	}
}
