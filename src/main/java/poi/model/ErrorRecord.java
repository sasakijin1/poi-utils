package poi.model;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @ClassName: ErrorRecord
 * @Description: 出错记录
 * @author wujinglei
 * @date 2014年6月11日 上午9:54:54
 *
 */
public class ErrorRecord {

	private final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	
	/**
	 * sheet序号
	 */
	private int sheetNo;
	
	/**
	 * 编码级错误
	 */
	private boolean codeError;
	
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
	private CellOptions cellOptions;
	
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
		String positionStr = "";
		if (sheetNo != 0 || row !=0 || cell != 0){
			positionStr += "表号:" + sheetNo + ",行号:" + row + "," + "列号:" + cell;
		}
		return sdf.format(recordTime) + ": " + positionStr 	+ "<br>信息:" + errorMsg + "<br>处理方式：" + handleType;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:34:56
	 * @Description:
	 */
	@SuppressWarnings("unused")
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
	public ErrorRecord(int sheetNo, int row, String errorMsg, String handleType,boolean codeError) {
		super();
		this.sheetNo = sheetNo;
		this.row = row;
		this.errorMsg = errorMsg;
		this.handleType = handleType;
		this.codeError = codeError;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午10:34:11
	 * @Description:
	 */
	public ErrorRecord(int sheetNo, int row, int cell, CellOptions cellOptions, String errorMsg, String handleType,boolean codeError) {
		this.sheetNo = sheetNo;
		this.row = row;
		this.cell = cell;
		this.cellOptions = cellOptions;
		this.errorMsg = errorMsg;
		this.handleType = handleType;
		this.codeError = codeError;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014年6月12日 上午11:05:04
	 * @Description:
	 */
	public ErrorRecord(int sheetNo, String errorMsg, String handleType,boolean codeError) {
		super();
		this.sheetNo = sheetNo;
		this.errorMsg = errorMsg;
		this.handleType = handleType;
		this.codeError = codeError;
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
	 * @return the cellOptions
	 */
	public CellOptions getCellOptions() {
		return cellOptions;
	}

	/**
	 * @param cellOptions the cellOptions to set
	 */
	public void setCellOptions(CellOptions cellOptions) {
		this.cellOptions = cellOptions;
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
