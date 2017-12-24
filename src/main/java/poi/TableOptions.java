package poi;

import java.util.List;

/**
 * @ClassName: TableOptions
 * @Description: 单个表格的设置
 * @author wujinglei
 * @date 2014-9-29 上午11:17:31
 *
 */
public final class TableOptions {
	
	public TableOptions(CellOptions[] cellOptions){
		this.cellOptions = cellOptions;
	}
	
	public TableOptions(CellOptions[] cellOptions,List exportData){
		this.cellOptions = cellOptions;
		this.exportData = exportData;
	}
	
	public TableOptions(CellOptions[] cellOptions,List exportData,int spacing){
		this.cellOptions = cellOptions;
		this.exportData = exportData;
		this.spacing = spacing;
	}
	
	/**
	 * 间距
	 */
	private int spacing = 0;

	/**
	 * 列设置
	 */
	private CellOptions[] cellOptions;
	
	/**
	 * 导出的数据
	 */
	private List exportData;
	
	/**
	 * 数据class类型
	 */
	private Class dataClazzType;
}
