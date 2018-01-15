package poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import poi.exception.XSSFCellTypeException;
import poi.model.CellDataType;
import poi.model.CellOptions;
import poi.model.ErrorRecord;
import poi.model.SheetOptions;
import poi.utils.BeanUtils;
import poi.utils.CellDataConverter;
import poi.utils.FieldUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.usermodel.CellType.STRING;

/**
 * The type Office io factory.
 *
 * @author wujinglei
 * @ClassName: OfficeIOFactory
 * @Description: OfficeIOFactory
 * @date 2014年6月11日 上午9:46:36
 */
public final class OfficeIoFactory {

    private final static Logger log = LoggerFactory.getLogger(OfficeIoFactory.class);

    /**
     * Export xlsx error record office io result.
     *
     * @param sheets        the sheets
     * @param errRecordRows the err record rows
     * @return office io result
     * @author: wujinglei
     * @date: 2014 -6-20 下午2:22:31
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
     * Export xlsx template office io result.
     *
     * @param sheetOptionsArray the sheet options array
     * @return office io result
     * @author: wujinglei
     * @date: 2014年6月12日 上午11:41:37
     * @Description: 导出模板
     */
    protected final OfficeIoResult exportXlsxTemplate(SheetOptions[] sheetOptionsArray) {
        // 实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheetOptionsArray);
        // 循环构建sheet
        for (int sheetIndex = 0; sheetIndex < sheetOptionsArray.length; sheetIndex++) {
            SheetOptions thisSheetOptions = sheetOptionsArray[sheetIndex];
            // 创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(thisSheetOptions.getSheetName());
            // 构建标题
            boolean hasSubTitle = buildTitle(result.getResultWorkbook(), sheet, thisSheetOptions.getCellOptions());

            createHideSelectSheet(result.getResultWorkbook(), thisSheetOptions, sheetIndex);

            // 导入DEMO数据
            buildDemoDataList(result.getResultWorkbook(), hasSubTitle, result, thisSheetOptions, sheet);
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
    private boolean buildTitle(Workbook workbook, Sheet sheet, CellOptions[] cells) {
        // 设置列头
        boolean hasSubTitle = buildTopTitle(workbook, sheet, cells, sheet.createRow(0));
        // 处理子列头
        if (hasSubTitle) {
            buildSubTitle(workbook, sheet, cells, sheet.createRow(1));
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
    private boolean buildTopTitle(Workbook workbook, Sheet sheet, CellOptions[] cells, Row titleRow) {
        boolean hasSubTitle = false;
        for (int titleIndex = 0, xlsCellIndex = 0; titleIndex < cells.length; titleIndex++) {
            CellOptions thisCellsOptions = cells[titleIndex];
            // 构建CELL
            Cell cell = createTitleCell(workbook, titleRow, xlsCellIndex, thisCellsOptions);

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
    private void buildSubTitle(Workbook workbook, Sheet sheet, CellOptions[] cells, Row subRow) {
        for (int titleIndex = 0, xlsCellIndex = 0; titleIndex < cells.length; titleIndex++) {
            if (cells[titleIndex].getSubCells() != null) {
                for (int subTitleIndex = 0; subTitleIndex < cells[titleIndex].getSubCells().length; subTitleIndex++) {
                    CellOptions thisCellsOptions = cells[titleIndex].getSubCells()[subTitleIndex];
                    Cell subTitleCell = createTitleCell(workbook, subRow, xlsCellIndex, thisCellsOptions);
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

    /**
     * @param workbook
     * @param hasSubTitle
     * @param result
     * @param thisSheetOptions
     * @param sheet
     */
    private void buildDemoDataList(Workbook workbook, boolean hasSubTitle, OfficeIoResult result, SheetOptions thisSheetOptions, Sheet sheet) {

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
                    Cell cell = createDataCell(workbook, row, xlsCellIndex, thisCellOptions);
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
                        Cell cell = createDataCell(workbook, row, xlsCellIndex, thisSubCellOptions);
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
    private long buildDataList(Workbook workbook, boolean hasSubTitle, SheetOptions thisSheetOptions, OfficeIoResult result, Sheet sheet, Integer sheetIndex) {

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
                        Cell cell = createDataCell(workbook, row, xlsCellIndex, thisCellOptions);
                        //写入内容
                        try {
                            setCellDataValue(sheet, cell, thisCellOptions, dataIndex + startRowIndex, xlsCellIndex, bean, dateStyle);
                        } catch (Exception e) {
                            recordSetCellDataValueException(result, row, sheetIndex, dataIndex, cellIndex, thisCellOptions, e);
                            continue rowLoop;
                        }
                        xlsCellIndex++;
                    } else {
                        for (int subIndex = 0; subIndex < cells[cellIndex].getSubCells().length; subIndex++) {
                            CellOptions thisSubCellOptions = cells[cellIndex].getSubCells()[subIndex];
                            //构建一个CELL
                            Cell cell = createDataCell(workbook, row, xlsCellIndex, thisSubCellOptions);
                            //写入内容
                            try {
                                setCellDataValue(sheet, cell, thisSubCellOptions, dataIndex + startRowIndex, xlsCellIndex, bean, dateStyle);
                            } catch (Exception e) {
                                recordSetCellDataValueException(result, row, sheetIndex, dataIndex, cellIndex, thisCellOptions, e);
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
     * Export xlsx office io result.
     *
     * @param sheetOptionsArray the sheet options array
     * @return office io result
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
            boolean hasSubTitle = buildTitle(result.getResultWorkbook(), sheet, cells);

            result.getResultTotal()[sheetIndex] = buildDataList(result.getResultWorkbook(), hasSubTitle, thisSheetOptions, result, sheet, sheetIndex);
        }

        return result;
    }

    /**
     * Import xlsx office io result.
     *
     * @param file   the file
     * @param sheets the sheets
     * @return office io result
     * @throws InvocationTargetException the invocation target exception
     * @throws IllegalAccessException    the illegal access exception
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:24:29
     * @Description: 导入XLSX
     */
    protected final OfficeIoResult importXlsx(File file, SheetOptions[] sheets) {
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
     * Import xlsx office io result.
     *
     * @param inputStream the input stream
     * @param sheets      the sheets
     * @return office io result
     * @throws InvocationTargetException the invocation target exception
     * @throws IllegalAccessException    the illegal access exception
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:24:29
     * @Description: 导入XLS
     */
    protected final OfficeIoResult importXlsx(InputStream inputStream, SheetOptions[] sheets) {
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
    private OfficeIoResult loadWorkbook(Workbook wb, SheetOptions[] sheets) {

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
            // reSet sheetSeq
            if (thisSheetOptions.getSheetSeq() == null) {
                thisSheetOptions.setSheetSeq(sheetIndex);
            }
            // 对每张表中的列进行读取处理
            CellOptions[] cells = thisSheetOptions.getCellOptions();

            // checkSkipRow
            if (thisSheetOptions.getSkipRows() == null) {
                thisSheetOptions.setSkipRows(1);
                for (CellOptions cellOptions : cells) {
                    if (cellOptions.getSubCells() != null) {
                        thisSheetOptions.setSkipRows(2);
                        break;
                    }
                }
            }
            if (thisSheetOptions.getSheetSeq() > sheetNumbers) {
                result.addErrorRecord(new ErrorRecord(thisSheetOptions.getSheetSeq(), "无法在文件中找到指定的sheet序号", "跳过sheet处理", true));
                continue;
            }

            // check entityDataType
//            if (thisSheetOptions.getDataClazzType() != null){
//                for (CellOptions cellOptions: cells){
//                    if (cellOptions.getSubCells() == null){
//                        if (cellOptions.getCellClass() == null){
//                            cellOptions.setCellClass(FieldUtils.getDeclaredFieldType(thisSheetOptions.getDataClazzType(),cellOptions.getKey()));
//                        }
//                    }else {
//                        for (CellOptions subCell: cellOptions.getSubCells()){
//                            subCell.setCellClass(FieldUtils.getDeclaredFieldType(thisSheetOptions.getDataClazzType(),subCell.getKey()));
//                        }
//                    }
//                }
//            }

            // 取提对应的sheet
            Sheet sheet = wb.getSheetAt(thisSheetOptions.getSheetSeq());
            List sheetList = new ArrayList();
            // 获取表中的总行数
            int rowsNum = sheet.getLastRowNum();
            //记录读取的总数
            result.setTotalRowCount(sheetIndex, (long) (rowsNum - thisSheetOptions.getSkipRows() + 1));

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
                if (activeRow != null) {
                    // 第一行的各列放在一个MAP中
                    Object resultObj;
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
                    int excelCellIndex = 0;
                    for (int cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                        Cell cell = activeRow.getCell(excelCellIndex);
                        if (cell != null) {
                            Object obj;
                            // 处理合并单元合问题
                            if (cells[cellIndex].getSubCells() != null) {
                                CellOptions[] subCells = cells[cellIndex].getSubCells();
                                for (int subCellIndex = 0; subCellIndex < subCells.length; subCellIndex++) {
                                    cell = activeRow.getCell(excelCellIndex);
                                    obj = getCellValue(cell, subCells[subCellIndex], wb);
                                    //判断规则
                                    if (!checkRule(cells, cellIndex + subCellIndex, obj, result, sheetIndex, activeRow)) {
                                        continue rowLoop;
                                    }
                                    setValueToObject(resultObj, subCells[subCellIndex], obj);
                                    excelCellIndex++;
                                }
                            } else {
                                obj = getCellValue(cell, cells[cellIndex], wb);
                                //判断规则
                                if (!checkRule(cells, cellIndex, obj, result, sheetIndex, activeRow)) {
                                    continue rowLoop;
                                }
                                setValueToObject(resultObj, cells[cellIndex], obj);
                                excelCellIndex++;
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

    private void setValueToObject (Object targetObj,CellOptions cellOptions,Object value){
        if (targetObj instanceof Map){
            ((Map) targetObj).put(cellOptions.getKey(),value);
        }else {
            BeanUtils.invokeSetter(targetObj, cellOptions.getKey(), value,cellOptions.getCellClass());
        }
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
    private Object getCellValue(Cell cell, CellOptions options, Workbook wb) {
        //如果有静态值，直接返回
        if (options != null && options.getHasStaticValue()) {
            return options.getStaticValue();
        }

        try {
            String cellValue;
            switch (cell.getCellTypeEnum()) {
                case BLANK:
                    cellValue = "";
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    cellValue = String.valueOf(cell.getCellFormula());
                    break;
                case NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        cellValue = CellDataConverter.date2Str(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()), CellDataConverter.DATE_FORMAT_DAY);
                    } else {
                        cellValue = CellDataConverter.scientificNotation(cellValue);
                    }
                    break;
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                default:
                    cellValue = null;
                    break;
            }
            //类型是否是自动匹配
            if (CellDataType.AUTO != options.getCellDataType()) {
                switch (options.getCellDataType()) {
                    case SELECT:
                        // TODO 后继处理
                        return "";
                    case VARCHAR:
                        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                            cellValue = CellDataConverter.matchNumber2Varchar(cellValue);
                        }
                        return cellValue;
                    case NUMBER:
                        try {
                            if (cellValue != null && !"".equals(cellValue)) {
                                return new BigDecimal(cellValue);
                            }
                        } catch (Exception e) {
                            return 0;
//                            TODO 添加异常记录
//                            throw new XSSFCellTypeException("Cell Type error,Can not read this CellValue: " + e.getMessage());
                        }
                    case DATE:
                        if (cellValue != null && !"".equals(cellValue)) {
                            return CellDataConverter.str2Date(cellValue);
                        }
                    case FORMULA:
                        if (CellType.FORMULA == cell.getCellTypeEnum()) {
                            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                            evaluator.evaluateFormulaCellEnum(cell);
                            return evaluator.evaluate(cell).getNumberValue();
                        } else {
                            return "";
//                            TODO 添加异常记录
//                            throw new XSSFCellTypeException("Cell Type error,Cell Type is not FORMULA: " + options.getKey());
                        }
                    default:
                        return "";
                }
            }

            return cellValue;
        } catch (Exception e) {
            e.printStackTrace();
            log.error("Cell Type error,cant read cell value: " + options.getKey());
            return "";
//            TODO 添加异常记录
//            throw new XSSFCellTypeException("Cell Type error,cant read cell value: " + options.getKey());
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
    private String getValue(CellOptions cellOptions, Object bean) {
        //如果有静态值，直接返回
        if (cellOptions.getHasStaticValue()) {
            return cellOptions.getStaticValue();
        }

        Object returnObj = null;
        returnObj = BeanUtils.invokeGetter(bean, cellOptions.getKey());
        if (cellOptions.getSelect()) {
//            return cellOptions.getCellSelectRealValue(returnObj.toString());
        }

        if (returnObj instanceof Date) {
            returnObj = CellDataConverter.date2Str((Date) returnObj, CellDataConverter.DATE_FORMAT_SEC);
        }

        //处理固定数据
        if (cellOptions.getFixedValue()) {
            returnObj = cellOptions.getFixedMap().get(String.valueOf(returnObj));
        }

        if (returnObj == null) {
            returnObj = "";
        }

        return String.valueOf(returnObj);
    }

    /**
     * @param workbook
     * @param row
     * @param xlsCellIndex
     * @param cellOptions
     * @return
     */
    private Cell createTitleCell(Workbook workbook, Row row, int xlsCellIndex, CellOptions cellOptions) {
        // 构建一个CELL
        Cell cell = row.createCell(xlsCellIndex);
        // 设置CELL为文本格式
        cell.setCellType(STRING);

        cell.setCellStyle(getCellStyle(workbook, cellOptions, true));

        return cell;
    }

    /**
     * @param workbook
     * @param row
     * @param xlsCellIndex
     * @param cellOptions
     * @return
     */
    private Cell createDataCell(Workbook workbook, Row row, int xlsCellIndex, CellOptions cellOptions) {
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
        cell.setCellStyle(getCellStyle(workbook, cellOptions, false));
        return cell;
    }

    /**
     * @param sheet
     * @param cell
     * @param cellOptions
     * @param rowIndex
     * @param xlsCellIndex
     * @param dataBean
     * @param dateStyle
     * @return
     * @throws IllegalAccessException
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     */
    private Cell setCellDataValue(Sheet sheet, Cell cell, CellOptions cellOptions, int rowIndex, int xlsCellIndex, Object dataBean, CellStyle dateStyle) {
        //写入内容
        if (cellOptions.getHasStaticValue()) {
            cell.setCellValue(cellOptions.getStaticValue());
        }
        if (cellOptions.getSelect()) {
            setSelectDataValidation(sheet,cellOptions.getKey() + "_TEXT",rowIndex,xlsCellIndex);
        }

        if (dataBean != null) {
            String reVal = getValue(cellOptions, dataBean);
            if (cellOptions.getCellDataType() == CellDataType.NUMBER && !StringUtils.isBlank(reVal)) {
                cell.setCellValue(new BigDecimal((reVal)).doubleValue());
            } else if (cellOptions.getCellDataType() == CellDataType.DATE) {
                try {
                    cell.setCellValue(CellDataConverter.str2Date(reVal));
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

    /**
     * @param workbook
     * @param cellOptions
     * @param isTitle
     * @return
     */
    private CellStyle getCellStyle(Workbook workbook, CellOptions cellOptions, boolean isTitle) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        try {
            if (!isTitle) {
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
        } catch (Exception e) {
            log.warn(e.getMessage());
        }
        return style;
    }

    /**
     * 统一处理异常
     *
     * @param result
     * @param row
     * @param sheetIndex
     * @param dataIndex
     * @param cellIndex
     * @param thisCellOptions
     * @param e
     */
    private void recordSetCellDataValueException(OfficeIoResult result, Row row, Integer sheetIndex, int dataIndex, int cellIndex, CellOptions thisCellOptions, Exception e) {
        try {
            throw e;
        } catch (IllegalArgumentException illegalArgumentException) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "数据异常(数据类型转换导致)", "跳过行处理", false));
            result.addErrorRecordRow(sheetIndex, row);
        } catch (NoSuchMethodException noSuchMethodException) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "属性异常(无法找到相应的属性)", "跳过行处理", true));
            result.addErrorRecordRow(sheetIndex, row);
        } catch (InvocationTargetException invocationTargetException) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "数据集异常(集合中的单个数据集异常)", "跳过行处理", true));
            result.addErrorRecordRow(sheetIndex, row);
        } catch (IllegalAccessException illegalAccessException) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "Bean方法调用异常(无法正常调用方法)", "跳过行处理", true));
            result.addErrorRecordRow(sheetIndex, row);
        } catch (Exception e1) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "Bean方法调用异常(无法正常调用方法)", "跳过行处理", true));
            result.addErrorRecordRow(sheetIndex, row);
        }
    }

    private void createHideSelectSheet(Workbook workbook, SheetOptions sheetOptions, int index) {
        Sheet selectTextSheet = workbook.createSheet("select" + "_" + index + "_Text");
        Sheet selectValueSheet = workbook.createSheet("select" + "_" + index + "_Value");

        CellOptions[] cellOptions = sheetOptions.getCellOptions();
        int selectRowIndex = 0;
        for (CellOptions thisCell : cellOptions) {
            if (thisCell.getSubCells() != null) {
                for (CellOptions subCell : thisCell.getSubCells()) {
                    if (subCell.getSelect()) {
                        setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, subCell);
                        createSelectNameList(selectTextSheet.getSheetName(), workbook, subCell.getKey() + "_TEXT", selectRowIndex, subCell.getSelectTextList().length, subCell.getSelectCascadeFlag());
                        selectRowIndex++;
                    }
                }
            } else {
                if (thisCell.getSelect()) {
                    setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, thisCell);
                    createSelectNameList(selectTextSheet.getSheetName(), workbook, thisCell.getKey() + "_TEXT", selectRowIndex, thisCell.getSelectTextList().length, thisCell.getSelectCascadeFlag());
                    selectRowIndex++;
                }
            }
        }

        workbook.setSheetHidden(workbook.getSheetIndex("select" + "_" + index + "_Text"), false);
        workbook.setSheetHidden(workbook.getSheetIndex("select" + "_" + index + "_Value"), false);
    }

    private void createSelectRow(Row currentRow, String[] textList) {
        if (textList != null && textList.length > 0) {
            int i = 0;
            for (String cellValue : textList) {
                Cell cell = currentRow.createCell(i++);
                cell.setCellValue(cellValue);
            }
        }
    }

    private void setSelectRow(Sheet selectTextSheet, Sheet selectValueSheet, int selectRowIndex, CellOptions cellOptions) {
        createSelectRow(selectTextSheet.createRow(selectRowIndex), cellOptions.getSelectTextList());
        createSelectRow(selectValueSheet.createRow(selectRowIndex), cellOptions.getSelectValueList());
    }

    private void createSelectNameList(String sheetName, Workbook workbook, String nameCode, int order, int size, boolean cascadeFlag) {
        Name name;
        name = workbook.createName();
        name.setNameName(nameCode);
        name.setRefersToFormula(sheetName + "!" + createSelectFormula(order + 1, size, cascadeFlag));
    }

    private static String createSelectFormula(int order, int size, boolean cascadeFlag) {
        char start = 'A';
        if (cascadeFlag) {
            start = 'B';
            if (size <= 25) {
                char end = (char) (start + size - 1);
                return "$" + start + "$" + order + ":$" + end + "$" + order;
            } else {
                char endPrefix = 'A';
                char endSuffix = 'A';
                //26-51之间，包括边界（仅两次字母表计算）
                if ((size - 25) / 26 == 0 || size == 51) {
                    //边界值
                    if ((size - 25) % 26 == 0) {
                        endSuffix = (char) ('A' + 25);
                    } else {
                        endSuffix = (char) ('A' + (size - 25) % 26 - 1);
                    }
                    //51以上
                } else {
                    if ((size - 25) % 26 == 0) {
                        endSuffix = (char) ('A' + 25);
                        endPrefix = (char) (endPrefix + (size - 25) / 26 - 1);
                    } else {
                        endSuffix = (char) ('A' + (size - 25) % 26 - 1);
                        endPrefix = (char) (endPrefix + (size - 25) / 26);
                    }
                }
                return "$" + start + "$" + order + ":$" + endPrefix + endSuffix + "$" + order;
            }
        } else {
            if (size <= 26) {
                char end = (char) (start + size - 1);
                return "$" + start + "$" + order + ":$" + end + "$" + order;
            } else {
                char endPrefix = 'A';
                char endSuffix = 'A';
                if (size % 26 == 0) {
                    endSuffix = (char) ('A' + 25);
                    if (size > 52 && size / 26 > 0) {
                        endPrefix = (char) (endPrefix + size / 26 - 2);
                    }
                } else {
                    endSuffix = (char) ('A' + size % 26 - 1);
                    if (size > 52 && size / 26 > 0) {
                        endPrefix = (char) (endPrefix + size / 26 - 1);
                    }
                }
                return "$" + start + "$" + order + ":$" + endPrefix + endSuffix + "$" + order;
            }
        }
    }

    private void setSelectDataValidation(Sheet sheet,String formulaString,int rowIndex,int xlsCellIndex) {
        sheet.addValidationData(getDataValidationByFormula(sheet,formulaString, rowIndex, xlsCellIndex));
    }

    private static DataValidation getDataValidationByFormula(Sheet sheet,String formulaString, int naturalRowIndex, int naturalColumnIndex) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
        XSSFDataValidationConstraint  dvConstraint = (XSSFDataValidationConstraint) dvHelper.createFormulaListConstraint(sheet.getWorkbook().getName(formulaString).getRefersToFormula());
        CellRangeAddressList addressList = new CellRangeAddressList(naturalRowIndex, naturalRowIndex, naturalColumnIndex, naturalColumnIndex);
        return dvHelper.createValidation(dvConstraint, addressList);
    }
}
