package poi;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.usermodel.CellType.STRING;

/**
 * @author wujinglei
 * @ClassName: OfficeIOFactory
 * @Description: OfficeIOFactory
 * @date 2014年6月11日 上午9:46:36
 */
public final class OfficeIoFactory {

    private final static Logger log = LoggerFactory.getLogger(OfficeIoFactory.class);

    private final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    private final SimpleDateFormat cstDateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);

    private SimpleDateFormat longDateFormat = new SimpleDateFormat("yyyy-mm-dd");

    private SimpleDateFormat shortDateFormat = new SimpleDateFormat("yyyy/mm/dd");

    private static final NumberFormat NUMBER_FORMAT = NumberFormat.getInstance();

    static {
        NUMBER_FORMAT.setGroupingUsed(false);
    }

    private final DecimalFormat decimalFormat = new DecimalFormat("0");


    /**
     * @param sheets
     * @param errRecordRows
     * @return
     * @author: wujinglei
     * @date: 2014-6-20 下午2:22:31
     * @Description: 导出errorRecord记录
     */
    protected final OfficeIoResult exportXlsxErrorRecord(SheetOptions[] sheets, Map<Integer, List> errRecordRows) {
        //实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheets);
        //循环构建sheet
        Set<Integer> keySet = errRecordRows.keySet();
        for (Integer index : keySet) {
            //创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(sheets[index].getSheetName());
            //取出CellOptions
            CellOptions[] cells = sheets[index].getCellOptions();
            //设置列头
            Row titleRow = sheet.createRow(0);
            for (int titleIndex = 0; titleIndex < cells.length; titleIndex++) {
                //构建一个CELL
                Cell titleCell = titleRow.createCell(titleIndex);
                //设置CELL为文本格式
                titleCell.setCellType(STRING);
                // 写入内容
                titleCell.setCellValue(cells[titleIndex].getColName());
            }
            //写入出错行记录
            List rowList = errRecordRows.get(index);
            int errorRowCount = rowList.size();
            for (int errorIndex = 0; errorIndex < errorRowCount; errorIndex++) {
                Row row = sheet.createRow(errorIndex + 1);
                if (rowList.get(errorIndex) instanceof Row) {
                    Iterator<Cell> it = ((Row) rowList.get(errorIndex)).cellIterator();
                    int cellIndex = 0;
                    while (it.hasNext()) {
                        Cell sourceCell = it.next();
                        Cell targetCell = row.createCell(cellIndex++);
                        targetCell.setCellType(STRING);
                        if (CellType.NUMERIC.equals(sourceCell.getCellTypeEnum())) {
                            targetCell.setCellValue(sourceCell.getNumericCellValue());
                        }
                        if (CellType.STRING.equals(sourceCell.getCellTypeEnum())) {
                            targetCell.setCellValue(sourceCell.getRichStringCellValue());
                        }
                        if (CellType.FORMULA.equals(sourceCell.getCellTypeEnum())) {
                            targetCell.setCellValue(sourceCell.getCellFormula());
                        }
                        if (CellType.BOOLEAN.equals(sourceCell.getCellTypeEnum())) {
                            targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        }
                        if (CellType.ERROR.equals(sourceCell.getCellTypeEnum())) {
                            targetCell.setCellValue(sourceCell.getErrorCellValue());
                        }
                    }
                } else if (rowList.get(errorIndex) instanceof String[]) {
                    String[] values = (String[]) rowList.get(errorIndex);
                    for (int i = 0; i < values.length; i++) {
                        Cell targetCell = row.createCell(i);
                        targetCell.setCellType(STRING);
                        targetCell.setCellValue(values[i]);
                    }
                }
            }
        }
        return result;
    }

    /**
     * @param sheetOptionsArray
     * @return
     * @author: wujinglei
     * @date: 2014年6月12日 上午11:41:37
     * @Description: 导出模板
     */
    protected final OfficeIoResult exportXlsxTempalet(SheetOptions[] sheetOptionsArray) {
        // 实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheetOptionsArray);
        // 循环构建sheet
        for (SheetOptions thisSheetOptions : sheetOptionsArray) {
            // 创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(thisSheetOptions.getSheetName());
            // 构建标题
            boolean hasSubTitle = buildTitle(result.getResultWorkbook(),sheet, thisSheetOptions.getCellOptions());
            // 导入DEMO数据
            buildDemoDataList(result.getResultWorkbook(),hasSubTitle, result, thisSheetOptions, sheet);
        }
        return result;
    }

    /**
     * 构建表头
     *
     * @param sheet
     * @param cells
     * @return
     */
    private boolean buildTitle(Workbook workbook,Sheet sheet, CellOptions[] cells) {
        // 设置列头
        boolean hasSubTitle = buildTopTitle(workbook,sheet, cells, sheet.createRow(0));
        // 处理子列头
        if (hasSubTitle) {
            buildSubTitle(workbook,sheet, cells, sheet.createRow(1));
        }
        return hasSubTitle;
    }

    /**
     * 构建顶部表头
     *
     * @param sheet
     * @param cells
     * @param titleRow
     * @return
     */
    private boolean buildTopTitle(Workbook workbook,Sheet sheet, CellOptions[] cells, Row titleRow) {
        boolean hasSubTitle = false;
        for (int titleIndex = 0, xlsCellIndex = 0; titleIndex < cells.length; titleIndex++) {
            CellOptions thisCellsOptions = cells[titleIndex];
            // 构建CELL
            Cell cell = createTitleCell(workbook,titleRow, xlsCellIndex, thisCellsOptions);

            cell.setCellValue(thisCellsOptions.getColName());

            if (thisCellsOptions.getSubCells() != null) {
                hasSubTitle = true;
                sheet.addMergedRegion(new CellRangeAddress(0, 0, xlsCellIndex, xlsCellIndex + thisCellsOptions.getSubCells().length - 1));
                xlsCellIndex += thisCellsOptions.getSubCells().length;
            } else {
                xlsCellIndex++;
            }
        }
        return hasSubTitle;
    }

    /**
     * 构建子表头
     *
     * @param sheet
     * @param cells
     * @param subRow
     */
    private void buildSubTitle(Workbook workbook,Sheet sheet, CellOptions[] cells, Row subRow) {
        for (int titleIndex = 0, xlsCellIndex = 0; titleIndex < cells.length; titleIndex++) {
            if (cells[titleIndex].getSubCells() != null) {
                for (int subTitleIndex = 0; subTitleIndex < cells[titleIndex].getSubCells().length; subTitleIndex++) {
                    CellOptions thisCellsOptions = cells[titleIndex].getSubCells()[subTitleIndex];
                    Cell subTitleCell = createTitleCell(workbook,subRow, xlsCellIndex, thisCellsOptions);
                    subTitleCell.setCellValue(thisCellsOptions.getColName());
                    xlsCellIndex++;
                }
            } else {
                CellRangeAddress region = new CellRangeAddress(0, 1, xlsCellIndex, xlsCellIndex);
                sheet.addMergedRegion(region);
                // 处理合并单元格的边框问题
                RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
                RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
                xlsCellIndex++;
            }
        }
    }

    private void buildDemoDataList(Workbook workbook,boolean hasSubTitle, OfficeIoResult result, SheetOptions thisSheetOptions, Sheet sheet) {

        CellStyle dateStyle = result.getResultWorkbook().createCellStyle();
        CreationHelper createHelper = result.getResultWorkbook().getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/mm/dd"));

        CellOptions[] cells = thisSheetOptions.getCellOptions();
        //循环新增每一条数据
        int startRowIndex = 1;
        if (hasSubTitle) {
            startRowIndex++;
        }

        for (int demoIndex = 0; demoIndex < 1; demoIndex++) {
            Row row = sheet.createRow(demoIndex + startRowIndex);
            //循环列配置为第一列赋值
            for (int cellIndex = 0, xlsCellIndex = 0; cellIndex < cells.length; cellIndex++) {
                CellOptions thisCellOptions = cells[cellIndex];
                if (thisCellOptions.getSubCells() == null) {
                    //构建一个CELL
                    Cell cell = createDataCell(workbook,row, xlsCellIndex, thisCellOptions);
                    try {
                        setCellDataValue(sheet, cell, thisCellOptions, demoIndex + startRowIndex, xlsCellIndex, null, dateStyle);
                    } catch (Exception e) {
                        log.warn(e.getMessage());
                        continue;
                    }
                    xlsCellIndex++;
                } else {
                    for (int subIndex = 0; subIndex < cells[cellIndex].getSubCells().length; subIndex++) {
                        CellOptions thisSubCellOptions = cells[cellIndex].getSubCells()[subIndex];
                        // 构建一个CELL
                        Cell cell = createDataCell(workbook,row, xlsCellIndex, thisSubCellOptions);
                        try {
                            setCellDataValue(sheet, cell, thisSubCellOptions, demoIndex + startRowIndex, xlsCellIndex, null, dateStyle);
                        } catch (Exception e) {
                            log.warn(e.getMessage());
                            continue;
                        }
                        xlsCellIndex++;
                    }
                }
            }
        }
    }

    /**
     * 构建数据内容
     *
     * @param hasSubTitle
     * @param thisSheetOptions
     * @param result
     * @param sheet
     * @param sheetIndex
     * @return
     */
    private long buildDataList(Workbook workbook,boolean hasSubTitle, SheetOptions thisSheetOptions, OfficeIoResult result, Sheet sheet, Integer sheetIndex) {

        CellStyle dateStyle = result.getResultWorkbook().createCellStyle();
        CreationHelper createHelper = result.getResultWorkbook().getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/mm/dd"));

        //取出当前sheet所要导出的数据
        List dataList = thisSheetOptions.getExportData();
        CellOptions[] cells = thisSheetOptions.getCellOptions();

        //循环新增每一条数据
        long successCount = 0;
        int startRowIndex = 1;
        if (hasSubTitle) {
            startRowIndex++;
        }

        if (dataList != null && dataList.size() > 0) {
            rowLoop:
            for (int dataIndex = 0; dataIndex < dataList.size(); dataIndex++) {
                //取出当前行的数据对象
                Object bean = dataList.get(dataIndex);
                //新增行
                Row row = sheet.createRow(dataIndex + startRowIndex);
                //循环列配置为第一列赋值
                for (int cellIndex = 0, xlsCellIndex = 0; cellIndex < cells.length; cellIndex++) {
                    CellOptions thisCellOptions = cells[cellIndex];

                    if (cells[cellIndex].getSubCells() == null) {
                        //构建一个CELL
                        Cell cell = createDataCell(workbook,row, xlsCellIndex, thisCellOptions);
                        //写入内容
                        try {
                            setCellDataValue(sheet, cell, thisCellOptions, dataIndex + startRowIndex, xlsCellIndex, bean, dateStyle);
                        } catch (Exception e) {
                            // TODO 统一处理异常
//                        try {
//                        } catch (IllegalArgumentException e) {
//                            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, cellOptions, "数据异常(数据类型转换导致)", "跳过行处理", false));
//                            result.addErrorRecordRow(sheetIndex, row);
//                            continue rowLoop;
//                        } catch (NoSuchMethodException e) {
//                            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, cellOptions, "属性异常(无法找到相应的属性)", "跳过行处理", true));
//                            result.addErrorRecordRow(sheetIndex, row);
//                            continue rowLoop;
//                        } catch (InvocationTargetException e) {
//                            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, cellOptions, "数据集异常(集合中的单个数据集异常)", "跳过行处理", true));
//                            result.addErrorRecordRow(sheetIndex, row);
//                            continue rowLoop;
//                        } catch (IllegalAccessException e) {
//                            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, cellOptions, "Bean方法调用异常(无法正常调用方法)", "跳过行处理", true));
//                            result.addErrorRecordRow(sheetIndex, row);
//                            continue rowLoop;
//                        }
                            continue rowLoop;
                        }
                        xlsCellIndex++;
                    } else {
                        for (int subIndex = 0; subIndex < cells[cellIndex].getSubCells().length; subIndex++) {
                            CellOptions thisSubCellOptions = cells[cellIndex].getSubCells()[subIndex];
                            //构建一个CELL
                            Cell cell = createDataCell(workbook,row, xlsCellIndex, thisSubCellOptions);
                            //写入内容
                            try {
                                setCellDataValue(sheet, cell, thisSubCellOptions, dataIndex + startRowIndex, xlsCellIndex, bean, dateStyle);
                            } catch (Exception e) {
                                // TODO 统一处理异常
                                continue rowLoop;
                            }
                            xlsCellIndex++;
                        }
                    }
                }
                //记录成功结果
                successCount++;
            }
        }

        return successCount;
    }

    /**
     * 导出版本为97-03版本的 未测试，不一定能用
     *
     * @param sheetOptionsArray
     * @return
     * @author wujinglei
     * @date:2016-11-04
     */
    @Deprecated
    protected final OfficeIoResult exportXls(SheetOptions[] sheetOptionsArray) {
        //实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheetOptionsArray);
        //循环构建sheet
        for (int sheetIndex = 0; sheetIndex < sheetOptionsArray.length; sheetIndex++) {
            SheetOptions thisSheetOptions = sheetOptionsArray[sheetIndex];
            CellOptions[] cells = thisSheetOptions.getCellOptions();
            //创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(thisSheetOptions.getSheetName());

            boolean hasSubTitle = buildTitle(result.getResultWorkbook(),sheet, cells);
            //将成功条数放入result中
            result.getResultTotal()[sheetIndex] = buildDataList(result.getResultWorkbook(),hasSubTitle, thisSheetOptions, result, sheet, sheetIndex);
        }

        return result;

    }

    /**
     * @param sheetOptionsArray
     * @return
     * @Description:导出
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:01:54
     */
    protected final OfficeIoResult exportXlsx(SheetOptions[] sheetOptionsArray) {
        //实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheetOptionsArray);
        //循环构建sheet
        for (int sheetIndex = 0; sheetIndex < sheetOptionsArray.length; sheetIndex++) {
            SheetOptions thisSheetOptions = sheetOptionsArray[sheetIndex];
            CellOptions[] cells = thisSheetOptions.getCellOptions();

            //创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(thisSheetOptions.getSheetName());
            boolean hasSubTitle = buildTitle(result.getResultWorkbook(),sheet, cells);

            result.getResultTotal()[sheetIndex] = buildDataList(result.getResultWorkbook(),hasSubTitle, thisSheetOptions, result, sheet, sheetIndex);
            ;
        }

        return result;
    }

    /**
     * @param file
     * @param sheets
     * @return
     * @throws Exception
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:24:29
     * @Description: 导入XLSX
     */
    protected final OfficeIoResult importXlsx(File file, SheetOptions[] sheets) throws InvocationTargetException, IllegalAccessException {
        // 按文件取出工作簿
        Workbook wb = null;
        try {
            wb = create(new FileInputStream(file));
        } catch (InvalidFormatException e) {
            log.error(e.getMessage());
        } catch (FileNotFoundException e) {
            log.error(e.getMessage());
        } catch (IOException e) {
            log.error(e.getMessage());
        }
        return loadWorkbook(wb, sheets);
    }

    /**
     * @param inputStream
     * @param sheets
     * @return
     * @throws Exception
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:24:29
     * @Description: 导入XLS
     */
    protected final OfficeIoResult importXlsx(InputStream inputStream, SheetOptions[] sheets) throws InvocationTargetException, IllegalAccessException {
        // 按文件取出工作簿
        Workbook wb = null;
        try {
            wb = create(inputStream);
        } catch (InvalidFormatException e) {
            log.error(e.getMessage());
        } catch (FileNotFoundException e) {
            log.error(e.getMessage());
        } catch (IOException e) {
            log.error(e.getMessage());
        }
        return loadWorkbook(wb, sheets);
    }

    /**
     * @param in
     * @return
     * @throws IOException
     * @author: wujinglei
     * @date: 2014年9月17日 14:20:00
     * @Description: 得到工作本
     */
    private Workbook create(InputStream in) throws IOException, InvalidFormatException {
        return WorkbookFactory.create(in);
    }

    /**
     * @param cells
     * @param cellIndex
     * @param obj
     * @param result
     * @param sheetIndex
     * @param activeRow
     * @return
     * @author: wujinglei
     * @date: 2014年7月8日 下午4:46:02
     * @Description: 判断规则
     */
    private Boolean checkRule(CellOptions[] cells, int cellIndex, Object obj, OfficeIoResult result, int sheetIndex, Row activeRow) {
        if (cells[cellIndex].getCellRule() != null) {
            switch (cells[cellIndex].getCellRule()) {
                case REQUIRED:
                    if (StringUtils.isBlank(String.valueOf(obj))) {
                        result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列不能为空", "跳过行处理", false));
                        result.addErrorRecordRow(sheetIndex, activeRow);
                        return cells[cellIndex].isKeepInput();
                    }
                    break;
                case EQUALSTO:
                    if (!cells[cellIndex].getCellRuleValue().equals(obj)) {
                        result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值"
                                + cells[cellIndex].getCellRuleValue() + "与读取出的值" + obj + "不相等", "跳过行处理", false));
                        result.addErrorRecordRow(sheetIndex, activeRow);
                        return cells[cellIndex].isKeepInput();
                    }
                case LONG:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Long.parseLong(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值不是长整型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return cells[cellIndex].isKeepInput();
                        }
                    }
                case INTEGER:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Integer.parseInt(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值不是整型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return cells[cellIndex].isKeepInput();
                        }
                    }
                case DOUBLE:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Double.parseDouble(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值不是浮点型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return cells[cellIndex].isKeepInput();
                        }
                    }
                case DATEFORMAT:
                    if (cells[cellIndex].getCellRuleValue() == null) {
                        result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值与所给的日期格式不相符", "跳过行处理", false));
                        result.addErrorRecordRow(sheetIndex, activeRow);
                        return cells[cellIndex].isKeepInput();
                    } else {
                        SimpleDateFormat cellSdf = new SimpleDateFormat(String.valueOf(cells[cellIndex].getCellRuleValue()));
                        try {
                            cellSdf.parse(String.valueOf(obj));
                        } catch (Exception e) {
                            result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值不是整型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return cells[cellIndex].isKeepInput();
                        }
                    }
            }
            return true;
        } else {
            return true;
        }
    }

    /**
     * @param wb
     * @param sheets
     * @return
     * @author: wujinglei
     * @date: 2014年6月11日 上午11:17:50
     * @Description: 按sheetOptions读取workbook中的数据
     */
    private OfficeIoResult loadWorkbook(Workbook wb, SheetOptions[] sheets) throws InvocationTargetException, IllegalAccessException {

        OfficeIoResult result = new OfficeIoResult(sheets);

        //文件异常时处理
        if (wb == null) {
            result.addErrorRecord(new ErrorRecord("文件无法读取或读取异常", "跳过所有处理", true));
            return result;
        }

        int sheetNumbers = wb.getNumberOfSheets();

        long successCount = 0;

        // 记录处理的数字
        result.setResultTotal(new Long[sheets.length]);
        result.setFileTotalRow(new Long[sheets.length]);

        for (int sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
            SheetOptions thisSheetOptions = sheets[sheetIndex];
            if (thisSheetOptions.getSheetSeq() > sheetNumbers) {
                result.addErrorRecord(new ErrorRecord(thisSheetOptions.getSheetSeq(), "无法在文件中找到指定的sheet序号", "跳过sheet处理", true));
                continue;
            }
            // 取提对应的sheeet
            Sheet sheet = wb.getSheetAt(thisSheetOptions.getSheetSeq());
            List sheetList = new ArrayList();
            // 获取表中的总行数
            int rowsNum = sheet.getLastRowNum();
            //记录读取的总数
            result.setTotalRowCount(sheetIndex, (long) (rowsNum - thisSheetOptions.getSkipRows() + 1));
            // 对每张表中的列进行读取处理
            CellOptions[] cells = thisSheetOptions.getCellOptions();
            // 循环每一行
            rowLoop:
            for (int row = 0; row <= rowsNum; row++) {
                //判断是否是在skipRow之内
                if (row < thisSheetOptions.getSkipRows()) {
                    continue;
                }
                // 取的当前行
                Row activeRow = sheet.getRow(row);
                // 判断当前行记录是否有有效
                if (activeRow != null && !(activeRow.equals(""))) {
                    // 第一行的各列放在一个MAP中
                    Object resultObj = null;
                    try {
                        resultObj = thisSheetOptions.getDataClazzType().newInstance();
                    } catch (InstantiationException e) {
                        log.error(e.getMessage());
                        resultObj = new HashMap();
                    } catch (IllegalAccessException e) {
                        log.error(e.getMessage());
                        resultObj = new HashMap();
                    }
                    // 循环每一列按列所给的参数进行处理
                    for (int cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                        Cell cell = activeRow.getCell(cellIndex);
                        if (cell != null) {
                            try {
                                Object obj = getCellValue(cell, cells[cellIndex], wb);
                                //判断规则
                                if (!checkRule(cells, cellIndex, obj, result, sheetIndex, activeRow)) {
                                    continue rowLoop;
                                }
                                BeanUtils.setProperty(resultObj, cells[cellIndex].getKey(), obj);
                            } catch (XSSFCellTypeException e) {
                                //列格式读取异常时，获得列名并抛出异常
                                result.addErrorRecord(new ErrorRecord(sheetIndex, row, cellIndex, cells[cellIndex], "当前列类型无法识别", "跳过行处理", false));
                                result.addErrorRecordRow(sheetIndex, activeRow);
                                if (!cells[cellIndex].isKeepInput()) {
                                    //暂时抛出异常，但不建议这么做。后期优化
                                    continue rowLoop;
                                } else {
                                    throw new GetCellValueRunTimeException(e.getMessage());
                                }
                            }
                        }
                    }
                    //将前当行所对应的MAP放入List中
                    sheetList.add(resultObj);
                } else {
                    result.addWrongRecord(new WrongRecord(sheetIndex, row, "导入的文件中空行数据", "跳过行处理", false));
                    continue;
                }
                //记录成功结果
                successCount++;
            }
            //将成功条数放入result中
            result.getResultTotal()[sheetIndex] = successCount;
            //将处理后的sheet的数据放入返回对象中
            result.addSheetList(sheetList);
        }
        return result;
    }


    /**
     * @param cell
     * @param options
     * @param wb
     * @return
     * @throws XSSFCellTypeException
     * @author: wujinglei
     * @date: 2014年6月11日 下午1:22:06
     * @Description: 按 options 取出列中的值
     */
    private Object getCellValue(Cell cell, CellOptions options, Workbook wb) throws XSSFCellTypeException {
        //如果有静态值，直接返回
        if (options != null && options.getHasStaticValue()) {
            return options.getStaticValue();
        }

        try {
            //类型是否是自动匹配
            if (CellDataType.AUTO != options.getCellDataType()) {
                switch (options.getCellDataType()) {
                    case SELECT:
                        String selectKey = "";
                        if (cell.getCellTypeEnum() == STRING) {
                            selectKey = cell.getStringCellValue();
                        } else {
                            if (cell.getCellTypeEnum() != CellType.NUMERIC) {
                                selectKey = decimalFormat.format(cell.getNumericCellValue());
                            }
                        }
                        return options.getCellSelectValue(selectKey);
                    case VARCHAR:
                        try {
                            // 如果数字类型先获取，在转换成字符串
                            if (cell.getCellTypeEnum() == STRING) {
                                return cell.getStringCellValue();
                            } else {
                                if (cell.getCellTypeEnum() != CellType.NUMERIC) {
                                    return decimalFormat.format(cell.getNumericCellValue());
                                }
                            }
                        } catch (Exception e) {
                            throw new XSSFCellTypeException("Cell Type error,Can not read this CellValue: " + e.getMessage());
                        }
                    case NUMBER:
                        try {
                            if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                BigDecimal varDou = new BigDecimal(cell.getNumericCellValue() + "");
                                if (varDou == null) {
                                    return "";
                                } else {
                                    return varDou;
                                }
                            } else {
                                String varStr = cell.getStringCellValue();
                                if (StringUtils.isBlank(varStr)) {
                                    return new BigDecimal(varStr);
                                } else {
                                    return "";
                                }
                            }
                        } catch (Exception e) {
                            throw new XSSFCellTypeException("Cell Type error,Can not read this CellValue: " + e.getMessage());
                        }
                    case DATE:
                        if (STRING == cell.getCellTypeEnum()) {
                            String cellDate = cell.getStringCellValue();
                            Date getcellDate = DateType(cellDate);
                            return StringUtils.isNotBlank(cell.getStringCellValue()) ? getcellDate : null;
                        } else {
                            Date varDate = cell.getDateCellValue();
                            return varDate;
                        }
                    case TIMESTAMP:
                        String varTimestamp = "";
                        try {
                            varTimestamp = cell.getStringCellValue();
                        } catch (IllegalStateException e) {
                            // 处理日期格式、时间格式
                            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                Date date = cell.getDateCellValue();
                                varTimestamp = simpleDateFormat.format(date);
                            } else if (cell.getCellStyle().getDataFormat() == 58) {
                                // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                                double dbValue = cell.getNumericCellValue();
                                Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(dbValue);
                                varTimestamp = simpleDateFormat.format(date);
                            }
                        }
                        if (StringUtils.isBlank(varTimestamp)) {
                            return "";
                        } else {
                            return varTimestamp;
                        }
                    case FORMULA:
                        if (CellType.FORMULA == cell.getCellTypeEnum()) {
                            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                            evaluator.evaluateFormulaCellEnum(cell);
                            CellValue cellValue = evaluator.evaluate(cell);
                            return cellValue.getNumberValue();
                        } else {
                            throw new XSSFCellTypeException("Cell Type error,Cell Type is not FORMULA: " + options.getKey());
                        }
                    default:
                        return "";
                }
            } else {
                switch (cell.getCellTypeEnum()) {
                    // 字符串
                    case STRING:
                        String value = cell.getStringCellValue();
                        if (StringUtils.isBlank(value)) {
                            value = "";
                        }
                        if (CellDataType.DATE == options.getCellDataType() && !StringUtils.isBlank(value)) {
                            return StringUtils.isNotBlank(cell.getStringCellValue()) ? longDateFormat.parse(cell.getStringCellValue()) : null;
                        } else {
                            return value;
                        }
                        // 数字
                    case NUMERIC:
                        // 处理日期格式、时间格式
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            Date date = cell.getDateCellValue();
                            return date;
                        } else if (cell.getCellStyle().getDataFormat() == 58) {
                            // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)  
                            double dbValue = cell.getNumericCellValue();
                            Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(dbValue);
                            return longDateFormat.format(date);
                        } else {
                            BigDecimal varDou = new BigDecimal(NUMBER_FORMAT.format(cell.getNumericCellValue()));
                            return varDou;
                        }
                        //工式
                    case FORMULA:
                        try {
                            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                            evaluator.evaluateFormulaCellEnum(cell);
                            CellValue cellValue = evaluator.evaluate(cell);
                            return cellValue.getNumberValue();
                        } catch (IllegalStateException e) {
                            throw new XSSFCellTypeException("Cell Type error,Cell Type is not FORMULA: " + options.getKey());
                        }
                        // 空值
                    case BLANK:
                        return "";
                    default:
                        throw new XSSFCellTypeException("Cell Type error,cant read cell value: " + options.getKey());
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            log.error("Cell Type error,cant read cell value: " + options.getKey());
            throw new XSSFCellTypeException("Cell Type error,cant read cell value: " + options.getKey());
        }
    }

    /**
     * @param cellOptions
     * @param bean
     * @return
     * @throws NoSuchMethodException
     * @throws SecurityException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws IllegalArgumentException
     * @author: wujinglei
     * @date: 2014年6月11日 下午4:47:09
     * @Description: 取出CELL所对应的值
     */
    private String getValue(CellOptions cellOptions, Object bean) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        //如果有静态值，直接返回
        if (cellOptions.getHasStaticValue()) {
            return cellOptions.getStaticValue();
        }

        Object returnObj = BeanUtils.getProperty(bean, cellOptions.getKey());

        if (cellOptions.getSelect()){
            return cellOptions.getCellSelectRealValue(returnObj.toString());
        }

        if (returnObj instanceof Date) {
            return simpleDateFormat.format(returnObj);
        }

        if (returnObj == null) {
            return "";
        }
        //处理固定数据
        if (cellOptions.getFixedValue()) {
            return (String) cellOptions.getFixedMap().get(String.valueOf(returnObj));
        }
        return String.valueOf(returnObj);
    }

    private Date DateType(String s) {
        Date celldate = null;
        try {
            celldate = longDateFormat.parse(s);
        } catch (Exception ex) {
            try {
                celldate = shortDateFormat.parse(s);
            } catch (ParseException e) {
            }
        }
        return celldate;
    }

    private Cell createTitleCell(Workbook workbook,Row row, int xlsCellIndex, CellOptions cellOptions) {
        // 构建一个CELL
        Cell cell = row.createCell(xlsCellIndex);
        // 设置CELL为文本格式
        cell.setCellType(STRING);

        cell.setCellStyle(getCellStyle(workbook,cellOptions,true));

        return cell;
    }

    private Cell createDataCell(Workbook workbook,Row row, int xlsCellIndex, CellOptions cellOptions) {
        // 构建一个CELL
        Cell cell = row.createCell(xlsCellIndex);
        // 设置CELL格式
        if (cellOptions.getCellDataType() != null) {
            switch (cellOptions.getCellDataType()) {
                case NUMBER:
                    cell.setCellType(CellType.NUMERIC);
                    break;
                default:
                    cell.setCellType(STRING);
                    break;
            }
        }
        cell.setCellStyle(getCellStyle(workbook,cellOptions,false));
        return cell;
    }

    private Cell setCellDataValue(Sheet sheet, Cell cell, CellOptions cellOptions, int rowIndex, int xlsCellIndex, Object dataBean, CellStyle dateStyle) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        //写入内容
        if (cellOptions.getHasStaticValue()) {
            cell.setCellValue(cellOptions.getStaticValue());
        } else if (cellOptions.getSelect()) {
            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
            XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(cellOptions.getSelectArray());
            CellRangeAddressList addressList = new CellRangeAddressList(rowIndex, rowIndex, xlsCellIndex, xlsCellIndex);
            XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);
            // 07默认setSuppressDropDownArrow(true);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);
        }

        if (dataBean != null) {
            String reVal = getValue(cellOptions, dataBean);
            if (cellOptions.getCellDataType() == CellDataType.NUMBER && !StringUtils.isBlank(reVal)) {
                cell.setCellValue(new BigDecimal((reVal)).doubleValue());
            } else if (cellOptions.getCellDataType() == CellDataType.DATE) {
                try {
                    cell.setCellValue(cstDateFormat.parse(reVal));
                } catch (ParseException e) {
                    log.warn(e.getMessage());
                    cell.setCellValue(reVal);
                }
                cell.setCellStyle(dateStyle);
            } else {
                cell.setCellValue(reVal);
            }
        }

        return cell;
    }

    private CellStyle getCellStyle(Workbook workbook,CellOptions cellOptions,boolean isTitle){
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        try{
            if (!isTitle){
                style.setFillForegroundColor(cellOptions.getCellStyleOptions().getDataForegroundColor());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setBorderRight(cellOptions.getCellStyleOptions().getDataBorder()[0]);
                style.setBorderTop(cellOptions.getCellStyleOptions().getDataBorder()[1]);
                style.setBorderLeft(cellOptions.getCellStyleOptions().getDataBorder()[2]);
                style.setBorderBottom(cellOptions.getCellStyleOptions().getDataBorder()[3]);
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setVerticalAlignment(VerticalAlignment.CENTER);

                font.setFontName(cellOptions.getCellStyleOptions().getDataFont());
                font.setColor(cellOptions.getCellStyleOptions().getDataFontColor());
                font.setFontHeightInPoints(cellOptions.getCellStyleOptions().getDataSize());
            } else {
                // TODO 处理XLS表头样式
                style.setFillForegroundColor(cellOptions.getCellStyleOptions().getTitleForegroundColor());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setBorderRight(cellOptions.getCellStyleOptions().getTitleBorder()[0]);
                style.setBorderTop(cellOptions.getCellStyleOptions().getTitleBorder()[1]);
                style.setBorderLeft(cellOptions.getCellStyleOptions().getTitleBorder()[2]);
                style.setBorderBottom(cellOptions.getCellStyleOptions().getTitleBorder()[3]);
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setVerticalAlignment(VerticalAlignment.CENTER);

                font.setFontName(cellOptions.getCellStyleOptions().getTitleFont());
                font.setColor(cellOptions.getCellStyleOptions().getTitleFontColor());
                font.setFontHeightInPoints(cellOptions.getCellStyleOptions().getTitleSize());
                style.setFont(font);
            }
        }catch (Exception e){
            log.warn(e.getMessage());
        }
        return style;
    }
}
