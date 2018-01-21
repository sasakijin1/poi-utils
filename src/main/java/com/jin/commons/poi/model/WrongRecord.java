package com.jin.commons.poi.model;

import com.jin.commons.poi.model.CellSettings;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @ClassName: WrongRecord
 * @Description: 警告记录
 * @author wujinglei
 * @date 2014年6月11日 上午9:54:54
 *
 */
public class WrongRecord {

	private final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

	/**
	 * sheet序号
	 */
	private int sheetNo;

	/**
	 * 编码级错误
	 */
	private boolean codeWrong;

	/**
	 * 行号
	 */
	private int row;

	/**
	 * 列号
	 */
	private int cell;

	/**
	 * 列参数
	 */
	private CellSettings cellSettings;

	/**
	 * 出错信息
	 */
	private String wrongMsg;

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
		String positionStr = "";
		if (sheetNo != 0 || row !=0 || cell != 0){
			positionStr += "表号:" + sheetNo + ",行号:" + row + "," + "列号:" + cell;
		}
		return sdf.format(recordTime) + ": " + positionStr 	+ "<br>信息:" + wrongMsg + "<br>处理方式：" + handleType;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:34:56
	 * @Description:
	 */
	@SuppressWarnings("unused")
	private WrongRecord(){

	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:55:48
	 * @Description:
	 */
	public WrongRecord(String wrongMsg, String handleType, boolean codeWrong) {
		this.wrongMsg = wrongMsg;
		this.handleType = handleType;
		this.codeWrong = codeWrong;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午11:08:42
	 * @Description:
	 */
	public WrongRecord(int sheetNo, int row, String wrongMsg, String handleType, boolean codeWrong) {
		super();
		this.sheetNo = sheetNo;
		this.row = row;
		this.wrongMsg = wrongMsg;
		this.handleType = handleType;
		this.codeWrong = codeWrong;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:34:11
	 * @Description:
	 */
	public WrongRecord(int sheetNo, int row, int cell, CellSettings cellSettings, String wrongMsg, String handleType, boolean codeWrong) {
		this.sheetNo = sheetNo;
		this.row = row;
		this.cell = cell;
		this.cellSettings = cellSettings;
		this.wrongMsg = wrongMsg;
		this.handleType = handleType;
		this.codeWrong = codeWrong;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午11:05:04
	 * @Description:
	 */
	public WrongRecord(int sheetNo, String wrongMsg, String handleType, boolean codeWrong) {
		super();
		this.sheetNo = sheetNo;
		this.wrongMsg = wrongMsg;
		this.handleType = handleType;
		this.codeWrong = codeWrong;
	}

	/**
	 * @return the sheetNo
	 */
	public int getSheetNo() {
		return sheetNo;
	}

	/**
	 * @param sheetNo the sheetNo to set
	 */
	public void setSheetNo(int sheetNo) {
		this.sheetNo = sheetNo;
	}

	/**
	 * @return the row
	 */
	public int getRow() {
		return row;
	}

	/**
	 * @param row the row to set
	 */
	public void setRow(int row) {
		this.row = row;
	}

	/**
	 * @return the cell
	 */
	public int getCell() {
		return cell;
	}

	/**
	 * @param cell the cell to set
	 */
	public void setCell(int cell) {
		this.cell = cell;
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
	 * @return the wrongMsg
	 */
	public String getWrongMsg() {
		return wrongMsg;
	}

	/**
	 * @param wrongMsg the wrongMsg to set
	 */
	public void setWrongMsg(String wrongMsg) {
		this.wrongMsg = wrongMsg;
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
	 * @return the codeWrong
	 */
	public boolean isCodeWrong() {
		return codeWrong;
	}

	/**
	 * @param codeWrong the codeWrong to set
	 */
	public void setCodeWrong(boolean codeWrong) {
		this.codeWrong = codeWrong;
	}
}
