package com.jin.commons.poi;

import com.jin.commons.poi.exception.SheetIndexException;
import com.jin.commons.poi.exception.XSSFCellTypeException;
import com.jin.commons.poi.model.*;
import com.jin.commons.poi.utils.BeanUtils;
import com.jin.commons.poi.utils.CellDataConverter;
import com.jin.commons.poi.utils.FieldUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

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
     * 导出异常数据记录
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
                titleCell.setCellType(CellType.STRING);
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
                        targetCell.setCellType(CellType.STRING);
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
                        targetCell.setCellType(CellType.STRING);
                        targetCell.setCellValue(values[i]);
                    }
                }
            }
        }
        return result;
    }

    /**
     * 导出XLSX模板
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
            SheetOptions thisSheetOptions;
            try {
                thisSheetOptions = checkCellOptions(result.getResultWorkbook(), sheetOptionsArray[sheetIndex], sheetIndex);
            } catch (SheetIndexException e) {
                result.addErrorRecord(new ErrorRecord(e.getMessage(), "跳过本SHEET所有处理", true));
                log.error(e.getMessage(),e);
                continue;
            }
            // 创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(thisSheetOptions.getSheetName());

            if (!StringUtils.isBlank(thisSheetOptions.getTitle())){
                buildTitle(sheet,thisSheetOptions);
            }

            // 构建标题
            boolean hasSubTitle = buildHeader(result.getResultWorkbook(), sheet, thisSheetOptions);

            createHideSelectSheet(result.getResultWorkbook(), thisSheetOptions, sheetIndex);

            // 导入DEMO数据
            buildDemoDataList(result.getResultWorkbook(), hasSubTitle, result, thisSheetOptions, sheet);
        }
        return result;
    }

    private void buildTitle(Sheet sheet,SheetOptions sheetOptions){
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(sheetOptions.getTitle());
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();

        style.setFillForegroundColor(sheetOptions.getTitleStyle().getTitleForegroundColor());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderRight(sheetOptions.getTitleStyle().getTitleBorder()[0]);
        style.setBorderTop(sheetOptions.getTitleStyle().getTitleBorder()[1]);
        style.setBorderLeft(sheetOptions.getTitleStyle().getTitleBorder()[2]);
        style.setBorderBottom(sheetOptions.getTitleStyle().getTitleBorder()[3]);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        font.setFontName(sheetOptions.getTitleStyle().getTitleFont());
        font.setColor(sheetOptions.getTitleStyle().getTitleFontColor());
        font.setFontHeightInPoints(sheetOptions.getTitleStyle().getTitleSize());
        style.setFont(font);

        titleCell.setCellStyle(style);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, sheetOptions.getCellCount() - 1);
        sheet.addMergedRegion(region);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
    }

    /**
     * 构建表头
     *
     * @param sheet
     * @param sheetOptions
     * @return
     */
    private boolean buildHeader(Workbook workbook, Sheet sheet, SheetOptions sheetOptions) {
        int startRow = 0;
        if (!StringUtils.isBlank(sheetOptions.getTitle())){
            startRow = 1;
        }
        // 设置列头
        boolean hasSubTitle = buildTopHeader(workbook, sheet, sheetOptions, sheet.createRow(startRow));
        // 处理子列头
        if (hasSubTitle) {
            buildSubHeader(workbook, sheet, sheetOptions, sheet.createRow(startRow + 1));
        }
        return hasSubTitle;
    }

    /**
     * 构建顶部表头
     *
     * @param sheet
     * @param sheetOptions
     * @param headerRow
     * @return
     */
    private boolean buildTopHeader(Workbook workbook, Sheet sheet, SheetOptions sheetOptions, Row headerRow) {
        boolean hasSubTitle = false;
        for (int titleIndex = 0, xlsCellIndex = 0; titleIndex < sheetOptions.getCellOptions().length; titleIndex++) {
            CellOptions thisCellsOptions = sheetOptions.getCellOptions()[titleIndex];
            // 构建CELL
            Cell cell = createHeaderCell(workbook, headerRow, xlsCellIndex, thisCellsOptions);

            sheetOptions.getCellAddressMap().put(thisCellsOptions.getKey(), CellReference.convertNumToColString(cell.getAddress().getColumn()));

            cell.setCellValue(thisCellsOptions.getColName());

            if (thisCellsOptions.getSubCells() != null) {
                hasSubTitle = true;
                sheet.addMergedRegion(new CellRangeAddress(headerRow.getRowNum(), headerRow.getRowNum(), xlsCellIndex, xlsCellIndex + thisCellsOptions.getSubCells().length - 1));
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
     * @param sheetOptions
     * @param subRow
     */
    private void buildSubHeader(Workbook workbook, Sheet sheet, SheetOptions sheetOptions, Row subRow) {
        for (int titleIndex = 0, xlsCellIndex = 0; titleIndex < sheetOptions.getCellOptions().length; titleIndex++) {
            CellOptions parentCellOptions = sheetOptions.getCellOptions()[titleIndex];
            if (parentCellOptions.getSubCells() != null) {
                for (int subTitleIndex = 0; subTitleIndex < parentCellOptions.getSubCells().length; subTitleIndex++) {
                    CellOptions thisCellsOptions = parentCellOptions.getSubCells()[subTitleIndex];
                    Cell subTitleCell = createHeaderCell(workbook, subRow, xlsCellIndex, thisCellsOptions);
                    sheetOptions.getCellAddressMap().put(thisCellsOptions.getKey(),CellReference.convertNumToColString(subTitleCell.getAddress().getColumn()));
                    subTitleCell.setCellValue(thisCellsOptions.getColName());
                    xlsCellIndex++;
                }
            } else {
                CellRangeAddress region = new CellRangeAddress(subRow.getRowNum() - 1, subRow.getRowNum(), xlsCellIndex, xlsCellIndex);
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
     * 创建DEMO数据
     * @param workbook
     * @param hasSubTitle
     * @param result
     * @param thisSheetOptions
     * @param sheet
     */
    private void buildDemoDataList(Workbook workbook, boolean hasSubTitle, OfficeIoResult result, SheetOptions thisSheetOptions, Sheet sheet) {

        CellOptions[] cells = thisSheetOptions.getCellOptions();
        //循环新增每一条数据
        int startRowIndex = 1;
        if (!StringUtils.isBlank(thisSheetOptions.getTitle())){
            startRowIndex += 1;
        }
        if (hasSubTitle) {
            startRowIndex += 1;
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
                        setCellDataValue(sheet, cell,thisSheetOptions, thisCellOptions, demoIndex + startRowIndex, xlsCellIndex, null);
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
                            setCellDataValue(sheet, cell,thisSheetOptions, thisSubCellOptions, demoIndex + startRowIndex, xlsCellIndex, null);
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

        //取出当前sheet所要导出的数据
        List dataList = thisSheetOptions.getExportData();
        CellOptions[] cells = thisSheetOptions.getCellOptions();

        //循环新增每一条数据
        long successCount = 0;
        int startRowIndex = 1;
        if (!StringUtils.isBlank(thisSheetOptions.getTitle())){
            startRowIndex += 1;
        }
        if (hasSubTitle) {
            startRowIndex += 1;
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
                            setCellDataValue(sheet, cell,thisSheetOptions, thisCellOptions, dataIndex + startRowIndex, xlsCellIndex, bean);
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
                                setCellDataValue(sheet, cell, thisSheetOptions,thisSubCellOptions, dataIndex + startRowIndex, xlsCellIndex, bean);
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
     * 导出XLSX
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
            SheetOptions thisSheetOptions;
            try {
                thisSheetOptions = checkCellOptions(result.getResultWorkbook(), sheetOptionsArray[sheetIndex], sheetIndex);
            } catch (SheetIndexException e) {
                result.addErrorRecord(new ErrorRecord(e.getMessage(), "跳过本SHEET所有处理", true));
                log.error(e.getMessage(),e);
                continue;
            }
            //创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(thisSheetOptions.getSheetName());

            if (!StringUtils.isBlank(thisSheetOptions.getTitle())){
                buildTitle(sheet,thisSheetOptions);
            }

            boolean hasSubTitle = buildHeader(result.getResultWorkbook(), sheet, thisSheetOptions);

            result.getResultTotal()[sheetIndex] = buildDataList(result.getResultWorkbook(), hasSubTitle, thisSheetOptions, result, sheet, sheetIndex);
        }

        return result;
    }

    /**
     * 导入XLSX
     *
     * @param file   the file
     * @param sheets the sheets
     * @return office io result
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:24:29
     * @Description: 导入XLSX
     */
    protected final OfficeIoResult importXlsx(File file, SheetOptions[] sheets) {
        // 按文件取出工作簿
        Workbook workbook = null;
        try {
            workbook = create(new FileInputStream(file));
        } catch (InvalidFormatException e) {
            log.error(e.getMessage());
        } catch (FileNotFoundException e) {
            log.error(e.getMessage());
        } catch (IOException e) {
            log.error(e.getMessage());
        }
        return loadWorkbook(workbook, sheets);
    }

    /**
     * Import xlsx office io result.
     *
     * @param inputStream the input stream
     * @param sheets      the sheets
     * @return office io result
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:24:29
     * @Description: 导入XLS
     */
    protected final OfficeIoResult importXlsx(InputStream inputStream, SheetOptions[] sheets) {
        // 按文件取出工作簿
        Workbook workbook = null;
        try {
            workbook = create(inputStream);
        } catch (InvalidFormatException e) {
            log.error(e.getMessage());
        } catch (FileNotFoundException e) {
            log.error(e.getMessage());
        } catch (IOException e) {
            log.error(e.getMessage());
        }
        return loadWorkbook(workbook, sheets);
    }

    /**
     * create Workbook
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
     * check cell Rules
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
                        return false;
                    }
                    break;
                case EQUALSTO:
                    if (!cells[cellIndex].getCellRuleValue().equals(obj)) {
                        result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值"
                                + cells[cellIndex].getCellRuleValue() + "与读取出的值" + obj + "不相等", "跳过行处理", false));
                        result.addErrorRecordRow(sheetIndex, activeRow);
                        return false;
                    }
                case LONG:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Long.parseLong(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值不是长整型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return false;
                        }
                    }
                case INTEGER:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Integer.parseInt(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值不是整型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return false;
                        }
                    }
                case DOUBLE:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Double.parseDouble(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值不是浮点型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return false;
                        }
                    }
                case DATEFORMAT:
                    if (cells[cellIndex].getCellRuleValue() == null) {
                        result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值与所给的日期格式不相符", "跳过行处理", false));
                        result.addErrorRecordRow(sheetIndex, activeRow);
                        return false;
                    } else {
                        SimpleDateFormat cellSdf = new SimpleDateFormat(String.valueOf(cells[cellIndex].getCellRuleValue()));
                        try {
                            cellSdf.parse(String.valueOf(obj));
                        } catch (Exception e) {
                            result.addErrorRecord(new ErrorRecord(sheetIndex, activeRow.getRowNum(), cellIndex, cells[cellIndex], "当前列预设值不是整型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return false;
                        }
                    }
                default:
                    break;
            }
        }
        return true;
    }

    /**
     * load Workwook data
     * @param workbook
     * @param sheets
     * @return
     * @author: wujinglei
     * @date: 2014年6月11日 上午11:17:50
     * @Description: 按sheetOptions读取workbook中的数据
     */
    private OfficeIoResult loadWorkbook(Workbook workbook, SheetOptions[] sheets) {

        OfficeIoResult result = new OfficeIoResult(sheets);

        //文件异常时处理
        if (workbook == null) {
            result.addErrorRecord(new ErrorRecord("文件无法读取或读取异常", "跳过所有处理", true));
            return result;
        }

        long successCount = 0;

        // 记录处理的数字
        result.setResultTotal(new Long[sheets.length]);
        result.setFileTotalRow(new Long[sheets.length]);

        for (int sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
            SheetOptions thisSheetOptions;
            try {
                thisSheetOptions = checkCellOptions(workbook, sheets[sheetIndex], sheetIndex);
            } catch (SheetIndexException e) {
                result.addErrorRecord(new ErrorRecord(e.getMessage(), "跳过本SHEET所有处理", true));
                log.error(e.getMessage(),e);
                continue;
            }

            CellOptions[] cells = thisSheetOptions.getCellOptions();

            // check selectSheet
            getSelectSheetMap(workbook,thisSheetOptions,sheetIndex);

            // 取提对应的sheet
            Sheet sheet = workbook.getSheetAt(thisSheetOptions.getSheetSeq());
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
                        if (thisSheetOptions.getDataClazzType() != null){
                            resultObj = thisSheetOptions.getDataClazzType().newInstance();
                        }else {
                            resultObj = new HashMap();
                        }
                    } catch (InstantiationException e) {
                        log.error(e.getMessage());
                        resultObj = new HashMap();
                    } catch (IllegalAccessException e) {
                        log.error(e.getMessage());
                        resultObj = new HashMap();
                    }
                    // 循环每一列按列所给的参数进行处理
                    int excelCellIndex = 0;
                    Map<String,String> selectTargetValueMap = new HashMap();
                    for (int cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                        Cell cell = activeRow.getCell(excelCellIndex);
                        if (cell != null) {
                            Object obj;
                            // 处理合并单元合问题
                            if (cells[cellIndex].getSubCells() != null) {
                                CellOptions[] subCells = cells[cellIndex].getSubCells();
                                for (int subCellIndex = 0; subCellIndex < subCells.length; subCellIndex++) {
                                    cell = activeRow.getCell(excelCellIndex);
                                    try {
                                        obj = getCellValue(cell, thisSheetOptions, subCells[subCellIndex], workbook,selectTargetValueMap);
                                    } catch (XSSFCellTypeException e) {
                                        recordSetCellDataValueException(result,activeRow,sheetIndex,row,cellIndex,subCells[subCellIndex],e);
                                        continue rowLoop;
                                    }
                                    //判断规则
                                    if (!checkRule(cells, cellIndex + subCellIndex, obj, result, sheetIndex, activeRow)) {
                                        continue rowLoop;
                                    }
                                    setValueToObject(resultObj, subCells[subCellIndex], obj);
                                    excelCellIndex++;
                                }
                            } else {
                                try {
                                    obj = getCellValue(cell, thisSheetOptions, cells[cellIndex], workbook,selectTargetValueMap);
                                } catch (XSSFCellTypeException e) {
                                    recordSetCellDataValueException(result,activeRow,sheetIndex,row,cellIndex,cells[cellIndex],e);
                                    continue rowLoop;
                                }
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

    /**
     * 将数据 放入对象
     * @param targetObj
     * @param cellOptions
     * @param value
     */
    private void setValueToObject (Object targetObj,CellOptions cellOptions,Object value){
        if (targetObj instanceof Map){
            ((Map) targetObj).put(cellOptions.getKey(),value);
        }else {
            BeanUtils.invokeSetter(targetObj, cellOptions.getKey(), value,cellOptions.getCellClass());
        }
    }

    /**
     * 读取单元格数据
     * @param cell
     * @param cellOptions
     * @param workbook
     * @return
     * @throws XSSFCellTypeException
     * @author: wujinglei
     * @date: 2014年6月11日 下午1:22:06
     * @Description: 按 options 取出列中的值
     */
    private Object getCellValue(Cell cell,SheetOptions sheetOptions, CellOptions cellOptions, Workbook workbook,Map<String,String> selectTargetValueMap) throws XSSFCellTypeException{
        //如果有静态值，直接返回
        String cellValue;
        try{
            if (cellOptions != null && cellOptions.getHasStaticValue()) {
                return cellOptions.getStaticValue();
            }

            switch (cell.getCellTypeEnum()) {
                case BLANK:
                    cellValue = null;
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
                        cellValue = CellDataConverter.date2Str(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()), DatePattern.DATE_FORMAT_DAY.getValue());
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
        }catch (Exception e){
            throw new XSSFCellTypeException("获取单元格数据时发生异常: " + e.getMessage());
        }

        if (sheetOptions.getSelectTargetSet().size() > 0){
            if (sheetOptions.getSelectTargetSet().contains(cellOptions.getKey())){
                selectTargetValueMap.put(cellOptions.getKey(),cellValue);
            }
        }

        // 处理下拉选择问题
        if (cellOptions.getSelect()){
            if (!cellOptions.getSelectCascadeFlag()){
                List<String> mapList = sheetOptions.getSelectMap().get(cellOptions.getKey() + "_TEXT");
                int matchIndex = mapList.indexOf(cellValue);
                if (matchIndex != -1){
                    cellValue = sheetOptions.getSelectMap().get(cellOptions.getKey() + "_TEXT_value").get(matchIndex);
                }else{
                    // TODO warn
                }
            }else {
                String formulaString = cellOptions.getKey() + "_" + selectTargetValueMap.get(cellOptions.getSelectTargetKey()) + "_TEXT";
                List<String> mapList = sheetOptions.getSelectMap().get(formulaString);
                int matchIndex = mapList.indexOf(cellValue);
                if (matchIndex != -1){
                    cellValue = sheetOptions.getSelectMap().get(formulaString + "_value").get(matchIndex);
                }else{
                    // TODO warn
                }
            }
        }

        //类型是否是自动匹配
        if (CellDataType.AUTO != cellOptions.getCellDataType() && cellValue != null) {
            switch (cellOptions.getCellDataType()) {
                case VARCHAR:
                    // XLS格式为数据的，去掉最后的.0
                    if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                        cellValue = CellDataConverter.matchNumber2Varchar(cellValue);
                    }
                    return cellValue;
                case NUMBER:
                    try {
                        if (!"".equals(cellValue)) {
                            if (cellOptions.getCellClass() == Double.class){
                                return Double.valueOf(cellValue);
                            }
                            if (cellOptions.getCellClass() == Float.class){
                                return Float.valueOf(cellValue);
                            }
                            return new BigDecimal(cellValue);
                        }
                    } catch (Exception e) {
                        throw new XSSFCellTypeException("Cell Value[" + cellValue + "] can not to Number: " + e.getMessage());
                    }
                case INTEGER:
                    try {
                        if (!"".equals(cellValue)) {
                            cellValue = CellDataConverter.matchNumber2Varchar(cellValue);
                            return Integer.valueOf(cellValue);
                        }
                        break;
                    } catch (Exception e) {
                        throw new XSSFCellTypeException("Cell Value[" + cellValue + "] can not to Integer: " + e.getMessage());
                    }
                case BIGINT:
                    try {
                        if (!"".equals(cellValue)) {
                            cellValue = CellDataConverter.matchNumber2Varchar(cellValue);
                            return Long.valueOf(cellValue);
                        }
                        break;
                    } catch (Exception e) {
                        throw new XSSFCellTypeException("Cell Value[" + cellValue + "] can not to Long: " + e.getMessage());
                    }
                case BOOLEAN:
                    try {
                        if (!"".equals(cellValue)) {
                            return Boolean.valueOf(cellValue);
                        }
                        break;
                    } catch (Exception e) {
                        throw new XSSFCellTypeException("Cell Value[" + cellValue + "] can not to Boolean: " + e.getMessage());
                    }
                case DATE:
                    try {
                        if (!"".equals(cellValue)) {
                            return CellDataConverter.str2Date(cellValue);
                        }
                        break;
                    } catch (Exception e) {
                        throw new XSSFCellTypeException("Cell Value[" + cellValue + "] can not to DATE: " + e.getMessage());
                    }
                case FORMULA:
                    if (CellType.FORMULA == cell.getCellTypeEnum()) {
                        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                        evaluator.evaluateFormulaCellEnum(cell);
                        return evaluator.evaluate(cell).getNumberValue();
                    } else {
                        throw new XSSFCellTypeException("Cell Type error,Cell Type is not FORMULA: " + cellOptions.getKey());
                    }
                default:
                    return null;
            }
        }

        return cellValue;

    }

    /**
     * 读取对象中的数据
     * @param cellOptions
     * @param bean
     * @return
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

        if (returnObj instanceof Date) {
            returnObj = CellDataConverter.date2Str((Date) returnObj, cellOptions.getPattern().getValue());
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
     * 创建 头
     * @param workbook
     * @param row
     * @param xlsCellIndex
     * @param cellOptions
     * @return
     */
    private Cell createHeaderCell(Workbook workbook, Row row, int xlsCellIndex, CellOptions cellOptions) {
        // 构建一个CELL
        Cell cell = row.createCell(xlsCellIndex);
        // 设置CELL为文本格式
        cell.setCellType(CellType.STRING);

        cell.setCellStyle(getCellStyle(workbook, cellOptions, true));

        return cell;
    }

    /**
     * 创建单元格
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
                    cell.setCellType(CellType.STRING);
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
     * @return
     */
    private void setCellDataValue(Sheet sheet, Cell cell, SheetOptions sheetOptions,CellOptions cellOptions, int rowIndex, int xlsCellIndex, Object dataBean) {
        //写入内容
        if (cellOptions.getHasStaticValue()) {
            cell.setCellValue(cellOptions.getStaticValue());
        }
        if (cellOptions.getSelect()) {
            StringBuilder formulaString = new StringBuilder();
            if (cellOptions.getSelectCascadeFlag()){
                String addressFlag = sheetOptions.getCellAddressMap().get(cellOptions.getSelectTargetKey());
                formulaString.append("INDIRECT(\"")
                        .append(cellOptions.getKey())
                        .append("_\"&")
                        .append(addressFlag)
                        .append(cell.getAddress().getRow() + 1)
                        .append("&\"_TEXT\")");
            }else {
                formulaString.append(cellOptions.getKey()).append("_TEXT");
            }
            setSelectDataValidation(sheet,formulaString.toString(),rowIndex,xlsCellIndex);
        }

        if (dataBean != null) {
            String reVal = getValue(cellOptions, dataBean);
            if (cellOptions.getCellDataType() == CellDataType.NUMBER && !StringUtils.isBlank(reVal)) {
                cell.setCellValue(new BigDecimal((reVal)).doubleValue());
            } else {
                cell.setCellValue(reVal);
            }
        }
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
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "数据异常(数据类型转换导致)", "跳过行处理:" + thisCellOptions.getKey(), false));
            result.addErrorRecordRow(sheetIndex, row);
        } catch (NoSuchMethodException noSuchMethodException) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "属性异常(无法找到相应的属性)", "跳过行处理:" + thisCellOptions.getKey(), true));
            result.addErrorRecordRow(sheetIndex, row);
        } catch (InvocationTargetException invocationTargetException) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "数据集异常(集合中的单个数据集异常)", "跳过行处理:" + thisCellOptions.getKey(), true));
            result.addErrorRecordRow(sheetIndex, row);
        } catch (IllegalAccessException illegalAccessException) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "Bean方法调用异常(无法正常调用方法)", "跳过行处理:" + thisCellOptions.getKey(), true));
            result.addErrorRecordRow(sheetIndex, row);
        } catch (Exception e1) {
            result.addErrorRecord(new ErrorRecord(sheetIndex, dataIndex, cellIndex, thisCellOptions, "Bean方法调用异常(无法正常调用方法)", "跳过行处理:" + thisCellOptions.getKey(), true));
            result.addErrorRecordRow(sheetIndex, row);
        }
    }

    /**
     * 处理下拉列表问题
     * @param workbook
     * @param sheetOptions
     * @param index
     */
    private void createHideSelectSheet(Workbook workbook, SheetOptions sheetOptions, int index) {
        Sheet selectTextSheet = workbook.createSheet("select" + "_" + index + "_text");
        Sheet selectValueSheet = workbook.createSheet("select" + "_" + index + "_value");

        CellOptions[] cellOptions = sheetOptions.getCellOptions();
        int selectRowIndex = 0;
        // 先处理没有联动的下拉
        Map<String,String[]> textMapping = new HashMap<String, String[]>();
        Map<String,String[]> valueMapping = new HashMap<String, String[]>();
        for (CellOptions thisCell : cellOptions) {
            if (thisCell.getSubCells() != null) {
                for (CellOptions subCell : thisCell.getSubCells()) {
                    if (subCell.getSelect() && !subCell.getSelectCascadeFlag()) {
                        setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, subCell.getSelectTextList(), subCell.getSelectValueList());
                        createSelectNameList(selectTextSheet.getSheetName(), workbook, subCell.getKey() + "_TEXT", selectRowIndex, subCell.getSelectTextList().length, subCell.getSelectCascadeFlag());
                        selectRowIndex++;
                        textMapping.put(subCell.getKey(),subCell.getSelectTextList());
                        valueMapping.put(subCell.getKey(),subCell.getSelectValueList());
                    }
                }
            } else {
                if (thisCell.getSelect() && !thisCell.getSelectCascadeFlag()) {
                    setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, thisCell.getSelectTextList(), thisCell.getSelectValueList());
                    createSelectNameList(selectTextSheet.getSheetName(), workbook, thisCell.getKey() + "_TEXT", selectRowIndex, thisCell.getSelectTextList().length, thisCell.getSelectCascadeFlag());
                    selectRowIndex++;
                    textMapping.put(thisCell.getKey(),thisCell.getSelectTextList());
                    valueMapping.put(thisCell.getKey(),thisCell.getSelectValueList());
                }
            }
        }

        // 处理有联动的下拉
        for (CellOptions thisCell : cellOptions) {
            if (thisCell.getSubCells() != null) {
                for (CellOptions subCell : thisCell.getSubCells()) {
                    if (subCell.getSelect() && subCell.getSelectCascadeFlag()) {
                        if (subCell.getSelectSourceList() != null && subCell.getSelectSourceList().size() > 0){
                            String[] targetTextArray = textMapping.get(subCell.getSelectTargetKey());
                            String[] targetValueArray = valueMapping.get(subCell.getSelectTargetKey());
                            if (targetTextArray != null && targetTextArray.length > 0){
                                if (targetValueArray != null && targetValueArray.length > 0) {
                                    Map<String, List<String>> addTextSelectMap = new HashMap<String, List<String>>();
                                    Map<String, List<String>> addValueSelectMap = new HashMap<String, List<String>>();
                                    for (String textKey : targetTextArray) {
                                        List<String> textList = new ArrayList<String>();
                                        textList.add(textKey);
                                        addTextSelectMap.put(textKey, textList);
                                    }
                                    for (String valueKey : targetValueArray) {
                                        List<String> valueList = new ArrayList<String>();
                                        valueList.add(valueKey);
                                        addValueSelectMap.put(valueKey, valueList);
                                    }
                                    for (Object obj : subCell.getSelectSourceList()) {
                                        int matchIndex = ArrayUtils.indexOf(targetValueArray, String.valueOf(BeanUtils.invokeGetter(obj, subCell.getBingKey())));
                                        if (matchIndex >= 0) {
                                            addTextSelectMap.get(targetTextArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, subCell.getCellSelect().getKey())));
                                            addValueSelectMap.get(targetValueArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, subCell.getCellSelect().getName())));
                                        }
                                    }

                                    for (Object obj : subCell.getSelectSourceList()) {
                                        int matchIndex = ArrayUtils.indexOf(targetValueArray, String.valueOf(BeanUtils.invokeGetter(obj, subCell.getBingKey())));
                                        if (matchIndex >= 0) {
                                            String[] addTextArray = new String[]{};
                                            addTextSelectMap.get(targetTextArray[matchIndex]).toArray(addTextArray);
                                            String[] addValueArray = new String[]{};
                                            addValueSelectMap.get(targetValueArray[matchIndex]).toArray(addValueArray);
                                            setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, addTextArray, addValueArray);
                                            setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, addTextArray, addValueArray);
                                            createSelectNameList(selectTextSheet.getSheetName(), workbook, subCell.getKey() + "_" + selectRowIndex + "_TEXT", selectRowIndex, addTextSelectMap.get(targetTextArray[matchIndex]).size(), subCell.getSelectCascadeFlag());
                                            selectRowIndex++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            } else {
                if (thisCell.getSelect() && thisCell.getSelectCascadeFlag()) {
                    if (thisCell.getSelectSourceList() != null && thisCell.getSelectSourceList().size() > 0){
                        String[] targetTextArray = textMapping.get(thisCell.getSelectTargetKey());
                        String[] targetValueArray = valueMapping.get(thisCell.getSelectTargetKey());
                        if (targetTextArray != null && targetTextArray.length > 0){
                            if (targetValueArray != null && targetValueArray.length > 0) {
                                LinkedHashMap<String, List<String>> addTextSelectMap = new LinkedHashMap<String, List<String>>();
                                LinkedHashMap<String, List<String>> addValueSelectMap = new LinkedHashMap<String, List<String>>();
                                Map<String, String> mappingMap = new HashMap();
                                for (int arraryIndex = 0; arraryIndex < targetTextArray.length; arraryIndex++) {
                                    mappingMap.put(targetTextArray[arraryIndex], targetValueArray[arraryIndex]);
                                }
                                for (String textKey : targetTextArray) {
                                    List<String> textList = new ArrayList<String>();
                                    textList.add(textKey);
                                    addTextSelectMap.put(textKey, textList);
                                }
                                for (String valueKey : targetValueArray) {
                                    List<String> valueList = new ArrayList<String>();
                                    valueList.add(valueKey);
                                    addValueSelectMap.put(valueKey, valueList);
                                }
                                for (Object obj : thisCell.getSelectSourceList()) {
                                    int matchIndex = ArrayUtils.indexOf(targetValueArray, String.valueOf(BeanUtils.invokeGetter(obj, thisCell.getBingKey())));
                                    if (matchIndex >= 0) {
                                        addTextSelectMap.get(targetTextArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, thisCell.getCellSelect().getName())));
                                        addValueSelectMap.get(targetValueArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, thisCell.getCellSelect().getKey())));
                                    }
                                }

                                for (String key : mappingMap.keySet()) {
                                    String[] addTextArray = new String[addTextSelectMap.get(key).size()];
                                    addTextSelectMap.get(key).toArray(addTextArray);
                                    String[] addValueArray = new String[addValueSelectMap.get(mappingMap.get(key)).size()];
                                    addValueSelectMap.get(mappingMap.get(key)).toArray(addValueArray);
                                    setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, addTextArray, addValueArray);
                                    createSelectNameList(selectTextSheet.getSheetName(), workbook, thisCell.getKey() + "_" + key + "_TEXT", selectRowIndex, addTextSelectMap.get(key).size(), thisCell.getSelectCascadeFlag());
                                    selectRowIndex++;
                                }
                            }
                        }
                    }
                }
            }
        }

        workbook.setSheetHidden(workbook.getSheetIndex("select" + "_" + index + "_text"), false);
        workbook.setSheetHidden(workbook.getSheetIndex("select" + "_" + index + "_value"), false);
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

    private void setSelectRow(Sheet selectTextSheet, Sheet selectValueSheet, int selectRowIndex, String[] textList, String[] valueList) {
        createSelectRow(selectTextSheet.createRow(selectRowIndex), textList);
        createSelectRow(selectValueSheet.createRow(selectRowIndex), valueList);
    }

    private void createSelectNameList(String sheetName, Workbook workbook, String nameCode, int order, int size, boolean cascadeFlag) {
        Name name;
        name = workbook.createName();
        name.setNameName(nameCode);
        if (cascadeFlag){
            size -= 1;
        }
        name.setRefersToFormula(sheetName + "!" + createSelectFormula(order + 1, size, cascadeFlag));
    }

    /**
     * 生成公式
     * @param order
     * @param size
     * @param cascadeFlag
     * @return
     */
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
        XSSFDataValidationConstraint  dvConstraint = new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.LIST,formulaString);
        CellRangeAddressList addressList = new CellRangeAddressList(rowIndex, rowIndex, xlsCellIndex, xlsCellIndex);
        DataValidation dataValidation = sheet.getDataValidationHelper().createValidation(dvConstraint, addressList);
        dataValidation.setShowErrorBox(true);
        sheet.addValidationData(dataValidation);
    }

    private SheetOptions checkCellOptions(Workbook workbook, SheetOptions sheetOptions,int sheetIndex) throws SheetIndexException {

        try {
            int sheetNumbers = workbook.getNumberOfSheets();

            // reSet sheetSeq
            if (sheetOptions.getSheetSeq() == null) {
                sheetOptions.setSheetSeq(sheetIndex);
            }
            CellOptions[] cells = sheetOptions.getCellOptions();

            // checkSkipRow
            if (sheetOptions.getSkipRows() == null) {
                sheetOptions.setSkipRows(1);
                for (CellOptions cellOptions : cells) {
                    if (cellOptions.getSubCells() != null) {
                        sheetOptions.setSkipRows(2);
                        break;
                    }
                }
            }

            if (!StringUtils.isBlank(sheetOptions.getTitle())){
                sheetOptions.setSkipRows(sheetOptions.getSkipRows() + 1);
            }

            // 处理联动下拉的Target问题
            int cellCount = 0;
            for (CellOptions cellOptions : cells) {
                if (cellOptions.getSubCells() != null) {
                    for(CellOptions subCellOptions: cellOptions.getSubCells()){
                        cellCount++;
                        if (subCellOptions.getSelectCascadeFlag()){
                            sheetOptions.getSelectTargetSet().add(subCellOptions.getSelectTargetKey());
                        }
                    }
                } else {
                    cellCount++;
                    if (cellOptions.getSelectCascadeFlag()){
                        sheetOptions.getSelectTargetSet().add(cellOptions.getSelectTargetKey());
                    }
                }
            }
            sheetOptions.setCellCount(cellCount);

            if (sheetOptions.getSheetSeq() > sheetNumbers) {
                throw new SheetIndexException("无法在文件中找到指定的sheet序号");
            }

            // check entityDataType
            if (sheetOptions.getDataClazzType() != null){
                for (CellOptions cellOptions: cells){
                    if (cellOptions.getSubCells() == null){
                        if (cellOptions.getCellClass() == null){
                            cellOptions.setCellClass(FieldUtils.getDeclaredFieldType(sheetOptions.getDataClazzType(),cellOptions.getKey()));
                            cellOptions.setCellDataType(FieldUtils.getCellDataType(cellOptions.getCellClass()));
                        }
                    }else {
                        for (CellOptions subCell: cellOptions.getSubCells()){
                            subCell.setCellClass(FieldUtils.getDeclaredFieldType(sheetOptions.getDataClazzType(),subCell.getKey()));
                            subCell.setCellDataType(FieldUtils.getCellDataType(subCell.getCellClass()));
                        }
                    }
                }
            }

            sheetOptions.setCellOptions(cells);
        } catch (Exception e){
            log.error(e.getMessage(), e);
            throw new SheetIndexException(e.getMessage());
        }
        return sheetOptions;
    }

    private void getSelectSheetMap(Workbook workbook,SheetOptions sheetOptions,int thisSheetIndex){
        List<Name> list = (List<Name>) workbook.getAllNames();
        for (Name name: list){
            if (name.getRefersToFormula().indexOf("_" + thisSheetIndex + "_") != 0){
                sheetOptions.getSelectMap().put(name.getNameName(),new ArrayList<String>());
                sheetOptions.getSelectMap().put(name.getNameName() + "_value",new ArrayList<String>());

                Sheet textSheet = workbook.getSheet(name.getRefersToFormula().split("!")[0]);
                Sheet valueSheet = workbook.getSheet(name.getRefersToFormula().split("!")[0].replace("_text","_value"));

                String address = name.getRefersToFormula().split("!")[1];

                int rowNum = Integer.valueOf(address.split(":")[0].substring(address.split(":")[0].lastIndexOf("$") + 1));
                String[] cellAddress = address.replaceAll("['$]","").replaceAll(String.valueOf(rowNum),"").split(":");

                Row textRow = textSheet.getRow(rowNum - 1);
                Row valueRow = valueSheet.getRow(rowNum - 1);

                for (int cellIndex = CellReference.convertColStringToIndex(cellAddress[0]); cellIndex <= CellReference.convertColStringToIndex(cellAddress[1]); cellIndex++) {
                    Cell textCell = textRow.getCell(cellIndex);
                    Cell valueCell = valueRow.getCell(cellIndex);
                    sheetOptions.getSelectMap().get(name.getNameName()).add(textCell.getStringCellValue());
                    sheetOptions.getSelectMap().get(name.getNameName() + "_value").add(valueCell.getStringCellValue());
                }
            }
        }
    }
}
