package com.jin.commons.poi;

import com.jin.commons.poi.exception.SheetIndexException;
import com.jin.commons.poi.exception.XSSFCellTypeException;
import com.jin.commons.poi.model.*;
import com.jin.commons.poi.utils.BeanUtils;
import com.jin.commons.poi.utils.CellDataConverter;
import com.jin.commons.poi.utils.DigestUtils;
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
    protected final OfficeIoResult exportXlsxErrorRecord(SheetSettings[] sheets, Map<Integer, List> errRecordRows) {
        //实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheets);
        //循环构建sheet
        Set<Integer> keySet = errRecordRows.keySet();
        for (Integer index : keySet) {
            try{

                SheetSettings thisSheetSettings = checkCellSettings(result.getResultWorkbook(), sheets[index], index);

                //创建sheet
                Sheet sheet = result.getResultWorkbook().createSheet(sheets[index].getSheetName());

                boolean hasSubTitle = buildHeader(result.getResultWorkbook(), sheet, thisSheetSettings);

                int startRow = hasSubTitle?1:0;
                //写入出错行记录
                List rowList = errRecordRows.get(index);
                int errorRowCount = rowList.size();
                for (int errorIndex = 0; errorIndex < errorRowCount; errorIndex++) {
                    Row row = sheet.createRow(errorIndex + 1 + startRow);
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
            }catch (Exception e){
                log.error(e.getMessage(),e);
            }
        }
        return result;
    }

    /**
     * 导出XLSX模板
     *
     * @param sheetSettingsArray the sheet settings array
     * @return office io result
     * @author: wujinglei
     * @date: 2014年6月12日 上午11:41:37
     * @Description: 导出模板
     */
    protected final OfficeIoResult exportXlsxTemplate(SheetSettings[] sheetSettingsArray) {
        // 实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheetSettingsArray);
        // 循环构建sheet
        for (int sheetIndex = 0; sheetIndex < sheetSettingsArray.length; sheetIndex++) {
            SheetSettings thisSheetSettings;
            try {
                thisSheetSettings = checkCellSettings(result.getResultWorkbook(), sheetSettingsArray[sheetIndex], sheetIndex);
            } catch (SheetIndexException e) {
                result.addErrorRecord(new ErrorRecord(e.getMessage(), "跳过本SHEET所有处理", true));
                log.error(e.getMessage(),e);
                continue;
            }
            // 创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(thisSheetSettings.getSheetName());

            if (!StringUtils.isBlank(thisSheetSettings.getTitle())){
                buildTitle(sheet,thisSheetSettings);
            }

            // 构建标题
            boolean hasSubTitle = buildHeader(result.getResultWorkbook(), sheet, thisSheetSettings);

            createHideSelectSheet(result.getResultWorkbook(), thisSheetSettings, sheetIndex);

            // 导入DEMO数据
            buildDemoDataList(result.getResultWorkbook(), hasSubTitle, result, thisSheetSettings, sheet);
        }
        return result;
    }

    /**
     * 构建标题
     * @param sheet
     * @param sheetSettings
     */
    private void buildTitle(Sheet sheet,SheetSettings sheetSettings){
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(sheetSettings.getTitle());
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();

        style.setFillForegroundColor(sheetSettings.getTitleStyle().getTitleForegroundColor());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderRight(sheetSettings.getTitleStyle().getTitleBorder()[0]);
        style.setBorderTop(sheetSettings.getTitleStyle().getTitleBorder()[1]);
        style.setBorderLeft(sheetSettings.getTitleStyle().getTitleBorder()[2]);
        style.setBorderBottom(sheetSettings.getTitleStyle().getTitleBorder()[3]);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        font.setFontName(sheetSettings.getTitleStyle().getTitleFont());
        font.setColor(sheetSettings.getTitleStyle().getTitleFontColor());
        font.setFontHeightInPoints(sheetSettings.getTitleStyle().getTitleSize());
        style.setFont(font);

        titleCell.setCellStyle(style);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, sheetSettings.getCellCount() - 1);
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
     * @param sheetSettings
     * @return
     */
    private boolean buildHeader(Workbook workbook, Sheet sheet, SheetSettings sheetSettings) {
        int startRow = 0;
        if (!StringUtils.isBlank(sheetSettings.getTitle())){
            startRow = 1;
        }
        // 设置列头
        boolean hasSubTitle = buildTopHeader(workbook, sheet, sheetSettings, sheet.createRow(startRow));
        // 处理子列头
        if (hasSubTitle) {
            buildSubHeader(workbook, sheet, sheetSettings, sheet.createRow(startRow + 1));
        }
        return hasSubTitle;
    }

    /**
     * 构建顶部表头
     *
     * @param sheet
     * @param sheetSettings
     * @param headerRow
     * @return
     */
    private boolean buildTopHeader(Workbook workbook, Sheet sheet, SheetSettings sheetSettings, Row headerRow) {
        boolean hasSubTitle = false;
        for (int titleIndex = 0, xlsCellIndex = 0; titleIndex < sheetSettings.getCellSettings().length; titleIndex++) {
            CellSettings thisCellsSettings = sheetSettings.getCellSettings()[titleIndex];
            // 构建CELL
            Cell cell = createHeaderCell(workbook, headerRow, xlsCellIndex, thisCellsSettings);

            cell.setCellValue(thisCellsSettings.getColName());

            if (thisCellsSettings.getSubCells() != null) {
                hasSubTitle = true;
                sheet.addMergedRegion(new CellRangeAddress(headerRow.getRowNum(), headerRow.getRowNum(), xlsCellIndex, xlsCellIndex + thisCellsSettings.getSubCells().length - 1));
                xlsCellIndex += thisCellsSettings.getSubCells().length;
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
     * @param sheetSettings
     * @param subRow
     */
    private void buildSubHeader(Workbook workbook, Sheet sheet, SheetSettings sheetSettings, Row subRow) {
        for (int titleIndex = 0, xlsCellIndex = 0; titleIndex < sheetSettings.getCellSettings().length; titleIndex++) {
            CellSettings parentCellSettings = sheetSettings.getCellSettings()[titleIndex];
            if (parentCellSettings.getSubCells() != null) {
                for (int subTitleIndex = 0; subTitleIndex < parentCellSettings.getSubCells().length; subTitleIndex++) {
                    CellSettings thisCellsSettings = parentCellSettings.getSubCells()[subTitleIndex];
                    Cell subTitleCell = createHeaderCell(workbook, subRow, xlsCellIndex, thisCellsSettings);
                    subTitleCell.setCellValue(thisCellsSettings.getColName());
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
     * @param thisSheetSettings
     * @param sheet
     */
    private void buildDemoDataList(Workbook workbook, boolean hasSubTitle, OfficeIoResult result, SheetSettings thisSheetSettings, Sheet sheet) {

        CellSettings[] cells = thisSheetSettings.getCellSettings();
        //循环新增每一条数据
        int startRowIndex = 1;
        if (!StringUtils.isBlank(thisSheetSettings.getTitle())){
            startRowIndex += 1;
        }
        if (hasSubTitle) {
            startRowIndex += 1;
        }

        for (int demoIndex = 0; demoIndex < 1; demoIndex++) {
            Row row = sheet.createRow(demoIndex + startRowIndex);
            //循环列配置为第一列赋值
            for (int cellIndex = 0, xlsCellIndex = 0; cellIndex < cells.length; cellIndex++) {
                CellSettings thisCellSettings = cells[cellIndex];
                if (thisCellSettings.getSubCells() == null) {
                    //构建一个CELL
                    Cell cell = createDataCell(workbook, row, xlsCellIndex, thisCellSettings);
                    try {
                        setCellDataValue(sheet, cell,thisSheetSettings, thisCellSettings, null);
                    } catch (Exception e) {
                        log.warn(e.getMessage());
                        continue;
                    }
                    xlsCellIndex++;
                } else {
                    for (int subIndex = 0; subIndex < cells[cellIndex].getSubCells().length; subIndex++) {
                        CellSettings thisSubCellSettings = cells[cellIndex].getSubCells()[subIndex];
                        // 构建一个CELL
                        Cell cell = createDataCell(workbook, row, xlsCellIndex, thisSubCellSettings);
                        try {
                            setCellDataValue(sheet, cell,thisSheetSettings, thisSubCellSettings, null);
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
     * @param thisSheetSettings
     * @param result
     * @param sheet
     * @param sheetIndex
     * @return
     */
    private long buildDataList(Workbook workbook, boolean hasSubTitle, SheetSettings thisSheetSettings, OfficeIoResult result, Sheet sheet, Integer sheetIndex) {

        //取出当前sheet所要导出的数据
        List dataList = thisSheetSettings.getExportData();
        CellSettings[] cells = thisSheetSettings.getCellSettings();

        //循环新增每一条数据
        long successCount = 0;
        int startRowIndex = 1;
        if (!StringUtils.isBlank(thisSheetSettings.getTitle())){
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
                    CellSettings thisCellSettings = cells[cellIndex];

                    if (cells[cellIndex].getSubCells() == null) {
                        //构建一个CELL
                        Cell cell = createDataCell(workbook, row, xlsCellIndex, thisCellSettings);
                        //写入内容
                        try {
                            setCellDataValue(sheet, cell,thisSheetSettings, thisCellSettings, bean);
                        } catch (Exception e) {
                            recordSetCellDataValueException(result, row, thisSheetSettings, cell.getAddress().formatAsString(), thisCellSettings, e);
                            continue rowLoop;
                        }
                        xlsCellIndex++;
                    } else {
                        for (int subIndex = 0; subIndex < cells[cellIndex].getSubCells().length; subIndex++) {
                            CellSettings thisSubCellSettings = cells[cellIndex].getSubCells()[subIndex];
                            //构建一个CELL
                            Cell cell = createDataCell(workbook, row, xlsCellIndex, thisSubCellSettings);
                            //写入内容
                            try {
                                setCellDataValue(sheet, cell, thisSheetSettings,thisSubCellSettings, bean);
                            } catch (Exception e) {
                                recordSetCellDataValueException(result, row, thisSheetSettings, cell.getAddress().formatAsString(), thisCellSettings, e);
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
     * @param sheetSettingsArray the sheet settings array
     * @return office io result
     * @Description:导出
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:01:54
     */
    protected final OfficeIoResult exportXlsx(SheetSettings[] sheetSettingsArray) {
        //实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheetSettingsArray);
        //循环构建sheet
        for (int sheetIndex = 0; sheetIndex < sheetSettingsArray.length; sheetIndex++) {
            SheetSettings thisSheetSettings;
            try {
                thisSheetSettings = checkCellSettings(result.getResultWorkbook(), sheetSettingsArray[sheetIndex], sheetIndex);
            } catch (SheetIndexException e) {
                result.addErrorRecord(new ErrorRecord(e.getMessage(), "跳过本SHEET所有处理", true));
                log.error(e.getMessage(),e);
                continue;
            }
            //创建sheet
            Sheet sheet = result.getResultWorkbook().createSheet(thisSheetSettings.getSheetName());

            if (!StringUtils.isBlank(thisSheetSettings.getTitle())){
                buildTitle(sheet,thisSheetSettings);
            }

            boolean hasSubTitle = buildHeader(result.getResultWorkbook(), sheet, thisSheetSettings);

            result.getResultTotal()[sheetIndex] = buildDataList(result.getResultWorkbook(), hasSubTitle, thisSheetSettings, result, sheet, sheetIndex);

            if (result.getErrors().size() > 0){
                result.setCompleted(false);
            }
        }
        cleanCacheData(sheetSettingsArray);

        result.setSheetSettings(sheetSettingsArray);

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
    protected final OfficeIoResult importXlsx(File file, SheetSettings[] sheets) {
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
    protected final OfficeIoResult importXlsx(InputStream inputStream, SheetSettings[] sheets) {
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
     * @param obj
     * @param result
     * @param sheetIndex
     * @param activeRow
     * @return
     * @author: wujinglei
     * @date: 2014年7月8日 下午4:46:02
     * @Description: 判断规则
     */
    private Boolean checkRule(CellSettings cellSettings, Cell cell, Object obj, OfficeIoResult result, int sheetIndex, Row activeRow) {
        if (cellSettings.getCellRule() != null) {
            switch (cellSettings.getCellRule()) {
                case REQUIRED:
                    if (obj == null || StringUtils.isBlank(String.valueOf(obj))) {
                        result.addErrorRecord(new ErrorRecord(cell.getSheet().getSheetName(), cell.getAddress().formatAsString(), cellSettings, "当前列不能为空", "跳过行处理", false));
                        result.addErrorRecordRow(sheetIndex, activeRow);
                        return false;
                    }
                    break;
                case EQUALSTO:
                    if (!cellSettings.getCellRuleValue().equals(obj)) {
                        result.addErrorRecord(new ErrorRecord(cell.getSheet().getSheetName(), cell.getAddress().formatAsString(), cellSettings, "当前列预设值"
                                + cellSettings.getCellRuleValue() + "与读取出的值" + obj + "不相等", "跳过行处理", false));
                        result.addErrorRecordRow(sheetIndex, activeRow);
                        return false;
                    }
                case LONG:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Long.parseLong(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(cell.getSheet().getSheetName(), cell.getAddress().formatAsString(), cellSettings, "当前列预设值不是长整型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return false;
                        }
                    }
                case INTEGER:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Integer.parseInt(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(cell.getSheet().getSheetName(), cell.getAddress().formatAsString(), cellSettings, "当前列预设值不是整型", "跳过行处理", false));
                            result.addErrorRecordRow(sheetIndex, activeRow);
                            return false;
                        }
                    }
                case DOUBLE:
                    if (StringUtils.isNotBlank(String.valueOf(obj))) {
                        try {
                            Double.parseDouble(String.valueOf(obj));
                        } catch (NumberFormatException ex) {
                            result.addErrorRecord(new ErrorRecord(cell.getSheet().getSheetName(), cell.getAddress().formatAsString(), cellSettings, "当前列预设值不是浮点型", "跳过行处理", false));
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
     * @Description: 按sheetSettings读取workbook中的数据
     */
    private OfficeIoResult loadWorkbook(Workbook workbook, SheetSettings[] sheets) {

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
            SheetSettings thisSheetSettings;
            try {
                thisSheetSettings = checkCellSettings(workbook, sheets[sheetIndex], sheetIndex);
            } catch (SheetIndexException e) {
                result.addErrorRecord(new ErrorRecord(e.getMessage(), "跳过本SHEET所有处理", true));
                log.error(e.getMessage(),e);
                continue;
            }

            CellSettings[] cells = thisSheetSettings.getCellSettings();

            // check selectSheet
            getSelectSheetMap(workbook,thisSheetSettings,sheetIndex);

            // 取提对应的sheet
            Sheet sheet = workbook.getSheetAt(thisSheetSettings.getSheetSeq());
            List sheetList = new ArrayList();
            // 获取表中的总行数
            int rowsNum = sheet.getLastRowNum();
            //记录读取的总数
            result.setTotalRowCount(sheetIndex, (long) (rowsNum - thisSheetSettings.getSkipRows() + 1));

            // 循环每一行
            rowLoop:
            for (int row = 0; row <= rowsNum; row++) {
                //判断是否是在skipRow之内
                if (row < thisSheetSettings.getSkipRows()) {
                    continue;
                }
                // 取的当前行
                Row activeRow = sheet.getRow(row);
                // 判断当前行记录是否有有效
                if (activeRow != null) {
                    // 第一行的各列放在一个MAP中
                    Object resultObj;
                    try {
                        if (thisSheetSettings.getDataClazzType() != null){
                            resultObj = thisSheetSettings.getDataClazzType().newInstance();
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
                                CellSettings[] subCells = cells[cellIndex].getSubCells();
                                for (int subCellIndex = 0; subCellIndex < subCells.length; subCellIndex++) {
                                    if (!subCells[subCellIndex].isSkip()){
                                        cell = activeRow.getCell(excelCellIndex);
                                        try {
                                            obj = getCellValue(cell, thisSheetSettings, subCells[subCellIndex], workbook,selectTargetValueMap);
                                        } catch (XSSFCellTypeException e) {
                                            recordSetCellDataValueException(result,activeRow,thisSheetSettings, cell.getAddress().formatAsString(),subCells[subCellIndex],e);
                                            continue rowLoop;
                                        }
                                        //判断规则
                                        if (!checkRule(subCells[subCellIndex], cell, obj, result, sheetIndex, activeRow)) {
                                            continue rowLoop;
                                        }
                                        setValueToObject(resultObj, subCells[subCellIndex], obj);
                                    }
                                    excelCellIndex++;
                                }
                            } else {
                                if (!cells[cellIndex].isSkip()){
                                    try {
                                        obj = getCellValue(cell, thisSheetSettings, cells[cellIndex], workbook,selectTargetValueMap);
                                    } catch (XSSFCellTypeException e) {
                                        recordSetCellDataValueException(result,activeRow,thisSheetSettings, cell.getAddress().formatAsString(),cells[cellIndex],e);
                                        continue rowLoop;
                                    }
                                    //判断规则
                                    if (!checkRule(cells[cellIndex], cell, obj, result, sheetIndex, activeRow)) {
                                        continue rowLoop;
                                    }
                                    setValueToObject(resultObj, cells[cellIndex], obj);
                                }
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

            if (result.getErrors().size() > 0){
                result.setCompleted(false);
            }
        }

        cleanCacheData(sheets);

        result.setSheetSettings(sheets);

        return result;
    }

    /**
     * 将数据 放入对象
     * @param targetObj
     * @param cellSettings
     * @param value
     */
    private void setValueToObject (Object targetObj,CellSettings cellSettings,Object value){
        if (targetObj instanceof Map){
            ((Map) targetObj).put(cellSettings.getKey(),value);
        }else {
            BeanUtils.invokeSetter(targetObj, cellSettings.getKey(), value,cellSettings.getCellClass());
        }
    }

    /**
     * 读取单元格数据
     * @param cell
     * @param cellSettings
     * @param workbook
     * @return
     * @throws XSSFCellTypeException
     * @author: wujinglei
     * @date: 2014年6月11日 下午1:22:06
     * @Description: 按 settings 取出列中的值
     */
    private Object getCellValue(Cell cell,SheetSettings sheetSettings, CellSettings cellSettings, Workbook workbook,Map<String,String> selectTargetValueMap) throws XSSFCellTypeException{
        //如果有静态值，直接返回
        String cellValue;
        try{
            if (cellSettings != null && cellSettings.getHasStaticValue()) {
                cellValue = cellSettings.getStaticValue();
            }else {
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
            }
        }catch (Exception e){
            throw new XSSFCellTypeException("获取单元格数据时发生异常: " + e.getMessage());
        }


        if (sheetSettings.getSelectTargetSet().size() > 0){
            if (sheetSettings.getSelectTargetSet().contains(cellSettings.getKey())){
                selectTargetValueMap.put(cellSettings.getKey(),cellValue);
            }
        }

        // 处理下拉选择问题
        if (cellSettings.getSelect()){
            if (!cellSettings.getSelectCascadeFlag()){
                String formulaString = cellSettings.getKey() + "_TEXT";
                List<String> mapList = sheetSettings.getSelectMap().get(DigestUtils.digestFormulaName(formulaString));
                int matchIndex = mapList.indexOf(cellValue);
                if (matchIndex != -1){
                    cellValue = sheetSettings.getSelectMap().get(DigestUtils.digestFormulaName(formulaString) + "_value").get(matchIndex);
                }else{
                    // TODO warn
                }
            }else {
                String formulaString = cellSettings.getKey() + "_" + selectTargetValueMap.get(cellSettings.getSelectTargetKey()) + "_TEXT";
                List<String> mapList = sheetSettings.getSelectMap().get(DigestUtils.digestFormulaName(formulaString));
                int matchIndex = mapList.indexOf(cellValue);
                if (matchIndex != -1){
                    cellValue = sheetSettings.getSelectMap().get(DigestUtils.digestFormulaName(formulaString) + "_value").get(matchIndex);
                }else{
                    // TODO warn
                }
            }
        }

        //类型是否是自动匹配
        if (CellDataType.AUTO != cellSettings.getCellDataType() && cellValue != null) {
            switch (cellSettings.getCellDataType()) {
                case VARCHAR:
                    // XLS格式为数据的，去掉最后的.0
                    if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                        cellValue = CellDataConverter.matchNumber2Varchar(cellValue);
                    }
                    return cellValue;
                case NUMBER:
                    try {
                        if (!"".equals(cellValue)) {
                            if (cellSettings.getCellClass() == Double.class){
                                return Double.valueOf(cellValue);
                            }
                            if (cellSettings.getCellClass() == Float.class){
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
                        throw new XSSFCellTypeException("Cell Type error,Cell Type is not FORMULA: " + cellSettings.getKey());
                    }
                default:
                    return null;
            }
        }

        return cellValue;

    }

    /**
     * 读取对象中的数据
     * @param cellSettings
     * @param bean
     * @return
     * @author: wujinglei
     * @date: 2014年6月11日 下午4:47:09
     * @Description: 取出CELL所对应的值
     */
    private String getValue(CellSettings cellSettings, Object bean) {
        //如果有静态值，直接返回
        if (cellSettings.getHasStaticValue()) {
            return cellSettings.getStaticValue();
        }

        Object returnObj = null;
        returnObj = BeanUtils.invokeGetter(bean, cellSettings.getKey());

        if (returnObj instanceof Date) {
            returnObj = CellDataConverter.date2Str((Date) returnObj, cellSettings.getPattern().getValue());
        }

        //处理固定数据
        if (cellSettings.getFixedValue()) {
            returnObj = cellSettings.getFixedMap().get(returnObj);
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
     * @param cellSettings
     * @return
     */
    private Cell createHeaderCell(Workbook workbook, Row row, int xlsCellIndex, CellSettings cellSettings) {
        // 构建一个CELL
        Cell cell = row.createCell(xlsCellIndex);
        // 设置CELL为文本格式
        cell.setCellType(CellType.STRING);

        cell.setCellStyle(getCellStyle(workbook, cellSettings, true));

        return cell;
    }

    /**
     * 创建单元格
     * @param workbook
     * @param row
     * @param xlsCellIndex
     * @param cellSettings
     * @return
     */
    private Cell createDataCell(Workbook workbook, Row row, int xlsCellIndex, CellSettings cellSettings) {
        // 构建一个CELL
        Cell cell = row.createCell(xlsCellIndex);
        // 设置CELL格式
        if (cellSettings.getCellDataType() != null) {
            switch (cellSettings.getCellDataType()) {
                case NUMBER:
                    cell.setCellType(CellType.NUMERIC);
                    break;
                default:
                    cell.setCellType(CellType.STRING);
                    break;
            }
        }
        cell.setCellStyle(getCellStyle(workbook, cellSettings, false));
        return cell;
    }

    /**
     * @param sheet
     * @param cell
     * @param cellSettings
     * @param dataBean
     * @return
     */
    private void setCellDataValue(Sheet sheet, Cell cell, SheetSettings sheetSettings,CellSettings cellSettings, Object dataBean) {
        //写入内容
        if (cellSettings.getHasStaticValue()) {
            cell.setCellValue(cellSettings.getStaticValue());
        }
        if (cellSettings.getSelect()) {
            StringBuilder formulaString = new StringBuilder();
            if (cellSettings.getSelectCascadeFlag()){
                String addressFlag = sheetSettings.getCellAddressMap().get(cellSettings.getSelectTargetKey());
                // =INDIRECT(VLOOKUP(A1,Sheet2!A:B,2,0))
                formulaString.append("INDIRECT(VLOOKUP(");
                formulaString.append("CONCATENATE(\"");
                formulaString.append(cellSettings.getKey());
                formulaString.append("_\",");
                formulaString.append(addressFlag);
                formulaString.append(cell.getAddress().getRow() + 1);
                formulaString.append("),select_");
                formulaString.append(sheetSettings.getSheetSeq());
                formulaString.append("_text");
                formulaString.append("!A:B,2,0))");
            }else {
                formulaString.append(DigestUtils.digestFormulaName(cellSettings.getKey() + "_TEXT"));
            }
            setSelectDataValidation(sheet,formulaString.toString(),cell.getRowIndex(),cell.getColumnIndex());
        }

        if (dataBean != null) {
            if (cellSettings.getCellDataType() != CellDataType.FORMULA){
                String reVal = getValue(cellSettings, dataBean);
                if (!cellSettings.getFixedValue()){
                    if (!StringUtils.isBlank(reVal)) {
                        switch (cellSettings.getCellDataType()){
                            case NUMBER:
                                cell.setCellValue(new BigDecimal((reVal)).doubleValue());
                                break;
                            case VARCHAR:
                                cell.setCellValue(reVal);
                                break;
                            case DATE:
                                cell.setCellValue(reVal);
                                break;
                            case BIGINT:
                                cell.setCellValue(Long.valueOf(reVal));
                                break;
                            case INTEGER:
                                cell.setCellValue(Integer.valueOf(reVal));
                                break;
                            case BOOLEAN:
                                cell.setCellValue(Boolean.valueOf(reVal));
                                break;
                            default:
                                cell.setCellValue(reVal);
                        }
                    }
                }else {
                    cell.setCellValue(reVal);
                }

            }else {
                FormulaSettings formulaSettings = cellSettings.getFormulaSettings();
                if (formulaSettings != null){
                    String formulaStr = createFormulaByGroup(sheetSettings,cell,formulaSettings.getFormulaType(),formulaSettings.getGroupName());
                    cell.setCellFormula(formulaStr);
                }

            }
        }
    }

    /**
     * @param workbook
     * @param cellSettings
     * @param isTitle
     * @return
     */
    private CellStyle getCellStyle(Workbook workbook, CellSettings cellSettings, boolean isTitle) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        try {
            if (!isTitle) {
                style.setFillForegroundColor(cellSettings.getCellStyleSettings().getDataForegroundColor());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setBorderRight(cellSettings.getCellStyleSettings().getDataBorder()[0]);
                style.setBorderTop(cellSettings.getCellStyleSettings().getDataBorder()[1]);
                style.setBorderLeft(cellSettings.getCellStyleSettings().getDataBorder()[2]);
                style.setBorderBottom(cellSettings.getCellStyleSettings().getDataBorder()[3]);
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setVerticalAlignment(VerticalAlignment.CENTER);

                font.setFontName(cellSettings.getCellStyleSettings().getDataFont());
                font.setColor(cellSettings.getCellStyleSettings().getDataFontColor());
                font.setFontHeightInPoints(cellSettings.getCellStyleSettings().getDataSize());
            } else {
                style.setFillForegroundColor(cellSettings.getCellStyleSettings().getTitleForegroundColor());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setBorderRight(cellSettings.getCellStyleSettings().getTitleBorder()[0]);
                style.setBorderTop(cellSettings.getCellStyleSettings().getTitleBorder()[1]);
                style.setBorderLeft(cellSettings.getCellStyleSettings().getTitleBorder()[2]);
                style.setBorderBottom(cellSettings.getCellStyleSettings().getTitleBorder()[3]);
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setVerticalAlignment(VerticalAlignment.CENTER);

                font.setFontName(cellSettings.getCellStyleSettings().getTitleFont());
                font.setColor(cellSettings.getCellStyleSettings().getTitleFontColor());
                font.setFontHeightInPoints(cellSettings.getCellStyleSettings().getTitleSize());
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
     * @param thisCellSettings
     * @param e
     */
    private void recordSetCellDataValueException(OfficeIoResult result, Row row, SheetSettings sheetSettings, String address, CellSettings thisCellSettings, Exception e) {
        try {
            throw e;
        } catch (IllegalArgumentException illegalArgumentException) {
            result.addErrorRecord(new ErrorRecord(sheetSettings.getSheetName(), address, thisCellSettings, "数据异常(数据类型转换导致)", "跳过行处理:" + thisCellSettings.getKey(), false));
            result.addErrorRecordRow(sheetSettings.getSheetSeq(), row);
        } catch (NoSuchMethodException noSuchMethodException) {
            result.addErrorRecord(new ErrorRecord(sheetSettings.getSheetName(), address, thisCellSettings, "属性异常(无法找到相应的属性)", "跳过行处理:" + thisCellSettings.getKey(), true));
            result.addErrorRecordRow(sheetSettings.getSheetSeq(), row);
        } catch (InvocationTargetException invocationTargetException) {
            result.addErrorRecord(new ErrorRecord(sheetSettings.getSheetName(), address, thisCellSettings, "数据集异常(集合中的单个数据集异常)", "跳过行处理:" + thisCellSettings.getKey(), true));
            result.addErrorRecordRow(sheetSettings.getSheetSeq(), row);
        } catch (IllegalAccessException illegalAccessException) {
            result.addErrorRecord(new ErrorRecord(sheetSettings.getSheetName(), address, thisCellSettings, "无法正常处理对应的数据：" + e.getMessage(), "跳过行处理:" + thisCellSettings.getKey(), true));
            result.addErrorRecordRow(sheetSettings.getSheetSeq(), row);
        } catch (Exception e1) {
            result.addErrorRecord(new ErrorRecord(sheetSettings.getSheetName(), address, thisCellSettings, "无法正常处理对应的数据：" + e.getMessage(), "跳过行处理:" + thisCellSettings.getKey(), true));
            result.addErrorRecordRow(sheetSettings.getSheetSeq(), row);
        }
    }

    /**
     * 处理下拉列表问题
     * @param workbook
     * @param sheetSettings
     * @param index
     */
    private void createHideSelectSheet(Workbook workbook, SheetSettings sheetSettings, int index) {
        Sheet selectTextSheet = workbook.createSheet("select" + "_" + index + "_text");
        Sheet selectValueSheet = workbook.createSheet("select" + "_" + index + "_value");

        CellSettings[] cellSettings = sheetSettings.getCellSettings();
        int selectRowIndex = 0;

        // 先处理没有联动的下拉
        Map<String,String[]> textMapping = new HashMap<String, String[]>();
        Map<String,String[]> valueMapping = new HashMap<String, String[]>();
        for (CellSettings thisCell : cellSettings) {
            if (thisCell.getSubCells() != null) {
                for (CellSettings subCell : thisCell.getSubCells()) {
                    if (subCell.getSelect() && !subCell.getSelectCascadeFlag()) {
                        setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, subCell.getSelectTextList(), subCell.getSelectValueList(),false);
                        createSelectNameList(selectTextSheet.getSheetName(), workbook, DigestUtils.digestFormulaName(subCell.getKey() + "_TEXT"), selectRowIndex, subCell.getSelectTextList().length, subCell.getSelectCascadeFlag());
                        selectRowIndex++;
                        textMapping.put(subCell.getKey(),subCell.getSelectTextList());
                        valueMapping.put(subCell.getKey(),subCell.getSelectValueList());
                    }
                }
            } else {
                if (thisCell.getSelect() && !thisCell.getSelectCascadeFlag()) {
                    setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, thisCell.getSelectTextList(), thisCell.getSelectValueList(),false);
                    createSelectNameList(selectTextSheet.getSheetName(), workbook, DigestUtils.digestFormulaName(thisCell.getKey() + "_TEXT"), selectRowIndex, thisCell.getSelectTextList().length, thisCell.getSelectCascadeFlag());
                    selectRowIndex++;
                    textMapping.put(thisCell.getKey(),thisCell.getSelectTextList());
                    valueMapping.put(thisCell.getKey(),thisCell.getSelectValueList());
                }
            }
        }

        // 处理有联动的下拉
        for (CellSettings thisCell : cellSettings) {
            if (thisCell.getSubCells() != null) {
                for (CellSettings subCell : thisCell.getSubCells()) {
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
                                            addTextSelectMap.get(targetTextArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, subCell.getCellSelectSettings().getKey())));
                                            addValueSelectMap.get(targetValueArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, subCell.getCellSelectSettings().getName())));
                                        }
                                    }

                                    for (Object obj : subCell.getSelectSourceList()) {
                                        int matchIndex = ArrayUtils.indexOf(targetValueArray, String.valueOf(BeanUtils.invokeGetter(obj, subCell.getBingKey())));
                                        if (matchIndex >= 0) {
                                            String formulaStr = DigestUtils.digestFormulaName(subCell.getKey() + "_" + selectRowIndex + "_TEXT");
                                            String[] addTextArray = new String[addTextSelectMap.get(targetTextArray[matchIndex]).size()];
                                            addTextSelectMap.get(targetTextArray[matchIndex]).toArray(addTextArray);
                                            String[] addValueArray = new String[addValueSelectMap.get(targetValueArray[matchIndex]).size()];
                                            addValueSelectMap.get(targetValueArray[matchIndex]).toArray(addValueArray);
                                            addTextArray[0] = subCell.getKey() + "_" + addTextArray[0];
                                            addTextArray = ArrayUtils.insert(1,addTextArray,formulaStr);
                                            addValueArray[0] = subCell.getKey() + "_" + addValueArray[0];
                                            addValueArray = ArrayUtils.insert(1,addValueArray,formulaStr);
                                            setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, addTextArray, addValueArray,true);
                                            createSelectNameList(selectTextSheet.getSheetName(), workbook, formulaStr, selectRowIndex, addTextSelectMap.get(targetTextArray[matchIndex]).size(), subCell.getSelectCascadeFlag());
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
                                        addTextSelectMap.get(targetTextArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, thisCell.getCellSelectSettings().getName())));
                                        addValueSelectMap.get(targetValueArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, thisCell.getCellSelectSettings().getKey())));
                                    }
                                }

                                for (String key : mappingMap.keySet()) {
                                    String formulaStr = DigestUtils.digestFormulaName(thisCell.getKey() + "_" + key + "_TEXT");
                                    String[] addTextArray = new String[addTextSelectMap.get(key).size()];
                                    addTextSelectMap.get(key).toArray(addTextArray);
                                    String[] addValueArray = new String[addValueSelectMap.get(mappingMap.get(key)).size()];
                                    addValueSelectMap.get(mappingMap.get(key)).toArray(addValueArray);
                                    addTextArray[0] = thisCell.getKey() + "_" + addTextArray[0];
                                    addTextArray = ArrayUtils.insert(1,addTextArray,formulaStr);
                                    addValueArray[0] = thisCell.getKey() + "_" + addValueArray[0];
                                    addValueArray = ArrayUtils.insert(1,addValueArray,formulaStr);
                                    setSelectRow(selectTextSheet, selectValueSheet, selectRowIndex, addTextArray, addValueArray,true);
                                    createSelectNameList(selectTextSheet.getSheetName(), workbook, formulaStr, selectRowIndex, addTextSelectMap.get(key).size(), thisCell.getSelectCascadeFlag());
                                    selectRowIndex++;
                                }
                            }
                        }
                    }
                }
            }
        }

        workbook.setSheetHidden(workbook.getSheetIndex("select" + "_" + index + "_text"), true);
        workbook.setSheetHidden(workbook.getSheetIndex("select" + "_" + index + "_value"), true);
    }

    /**
     * 创建选择下拉行数据
     * @param currentRow
     * @param textList
     */
    private void createSelectRow(Row currentRow, String[] textList, boolean cascadeFlag) {
        if (textList != null && textList.length > 0) {
            int i = 2;
            if (cascadeFlag){
                i = 0;
            }
            for (String cellValue : textList) {
                Cell cell = currentRow.createCell(i++);
                cell.setCellValue(cellValue);
            }
        }
    }

    /**
     * 设置下拉列表行数据
     * @param selectTextSheet
     * @param selectValueSheet
     * @param selectRowIndex
     * @param textList
     * @param valueList
     */
    private void setSelectRow(Sheet selectTextSheet, Sheet selectValueSheet, int selectRowIndex, String[] textList, String[] valueList, boolean cascadeFlag) {
        createSelectRow(selectTextSheet.createRow(selectRowIndex), textList, cascadeFlag);
        createSelectRow(selectValueSheet.createRow(selectRowIndex), valueList, cascadeFlag);
    }

    /**
     * 构建函数名
     * @param sheetName
     * @param workbook
     * @param nameCode
     * @param order
     * @param size
     * @param cascadeFlag
     */
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
        char start = 'C';
        if (cascadeFlag) {
            if (size == 0){
                return "$" + start + "$" + order;
            }
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
            if (size == 0){
                return "$" + start + "$" + order;
            }
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

    /**
     * 设置下拉校验规则
     * @param sheet
     * @param formulaString
     * @param rowIndex
     * @param xlsCellIndex
     */
    private void setSelectDataValidation(Sheet sheet,String formulaString,int rowIndex,int xlsCellIndex) {
        XSSFDataValidationConstraint  dvConstraint = new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.LIST,formulaString);
        CellRangeAddressList addressList = new CellRangeAddressList(rowIndex, rowIndex, xlsCellIndex, xlsCellIndex);
        DataValidation dataValidation = sheet.getDataValidationHelper().createValidation(dvConstraint, addressList);
        dataValidation.setShowErrorBox(true);
        sheet.addValidationData(dataValidation);
    }

    /**
     * 检查Cell信息
     * @param workbook
     * @param sheetSettings
     * @param sheetIndex
     * @return
     * @throws SheetIndexException
     */
    private SheetSettings checkCellSettings(Workbook workbook, SheetSettings sheetSettings,int sheetIndex) throws SheetIndexException {

        try {
            int sheetNumbers = workbook.getNumberOfSheets();

            // reSet sheetSeq
            if (sheetSettings.getSheetSeq() == null) {
                sheetSettings.setSheetSeq(sheetIndex);
            }
            CellSettings[] cells = sheetSettings.getCellSettings();

            // checkSkipRow
            if (sheetSettings.getSkipRows() == null) {
                sheetSettings.setSkipRows(1);
                for (CellSettings cellSettings : cells) {
                    if (cellSettings.getSubCells() != null) {
                        sheetSettings.setSkipRows(2);
                        break;
                    }
                }
            }

            if (!StringUtils.isBlank(sheetSettings.getTitle())){
                sheetSettings.setSkipRows(sheetSettings.getSkipRows() + 1);
            }

            // 处理联动下拉的Target问题
            int cellCount = 0;
            Map<String,Set<String>> formulaMap = new HashMap<String, Set<String>>();
            for (CellSettings cellSettings : cells) {
                if (cellSettings.getSubCells() != null) {
                    for(CellSettings subCellSettings: cellSettings.getSubCells()){
                        sheetSettings.getCellAddressMap().put(subCellSettings.getKey(),CellReference.convertNumToColString(cellCount));
                        if (subCellSettings.getCellDataType() != CellDataType.FORMULA){
                            if (subCellSettings.getFormulaGroupNames() != null){
                                for(String groupName: subCellSettings.getFormulaGroupNames()){
                                    formulaMap.put(groupName,setformulaGroupName(formulaMap.get(groupName),subCellSettings.getKey()));
                                }
                            }
                        }

                        if (subCellSettings.getSelectCascadeFlag()){
                            sheetSettings.getSelectTargetSet().add(subCellSettings.getSelectTargetKey());
                        }
                        cellCount++;
                    }
                } else {
                    sheetSettings.getCellAddressMap().put(cellSettings.getKey(), CellReference.convertNumToColString(cellCount));
                    if (cellSettings.getCellDataType() != CellDataType.FORMULA){
                        if (cellSettings.getFormulaGroupNames() != null){
                            for(String groupName: cellSettings.getFormulaGroupNames()){
                                formulaMap.put(groupName,setformulaGroupName(formulaMap.get(groupName),cellSettings.getKey()));
                            }
                        }
                    }

                    if (cellSettings.getSelectCascadeFlag()){
                        sheetSettings.getSelectTargetSet().add(cellSettings.getSelectTargetKey());
                    }
                    cellCount++;
                }
            }
            sheetSettings.setCellCount(cellCount);

            for (CellSettings cellSettings : cells) {
                if (cellSettings.getSubCells() != null) {
                    for (CellSettings subCellSettings : cellSettings.getSubCells()) {
                        if (subCellSettings.getCellDataType() == CellDataType.FORMULA){
                            subCellSettings.getFormulaSettings().setGroupName(formulaMap.get(subCellSettings.getFormulaGroupNames()[0]));
                        }
                    }
                }else {
                    if (cellSettings.getCellDataType() == CellDataType.FORMULA){
                        cellSettings.getFormulaSettings().setGroupName(formulaMap.get(cellSettings.getFormulaGroupNames()[0]));
                    }
                }
            }

            if (sheetSettings.getSheetSeq() > sheetNumbers) {
                throw new SheetIndexException("无法在文件中找到指定的sheet序号");
            }

            // check entityDataType
            if (sheetSettings.getDataClazzType() != null){
                for (CellSettings cellSettings: cells){
                    if (cellSettings.getSubCells() == null){
                        if (cellSettings.getCellDataType() != CellDataType.FORMULA){
                            if (cellSettings.getCellClass() == null){
                                cellSettings.setCellClass(FieldUtils.getDeclaredFieldType(sheetSettings.getDataClazzType(),cellSettings.getKey()));
                                cellSettings.setCellDataType(FieldUtils.getCellDataType(cellSettings.getCellClass()));
                            }
                        }else {
                            continue;
                        }
                    }else {
                        for (CellSettings subCell: cellSettings.getSubCells()){
                            if (subCell.getCellDataType() != CellDataType.FORMULA){
                                subCell.setCellClass(FieldUtils.getDeclaredFieldType(sheetSettings.getDataClazzType(),subCell.getKey()));
                                subCell.setCellDataType(FieldUtils.getCellDataType(subCell.getCellClass()));
                            }else {
                                continue;
                            }
                        }
                    }
                }
            }

            sheetSettings.setCellSettings(cells);
        } catch (Exception e){
            log.error(e.getMessage(), e);
            throw new SheetIndexException(e.getMessage());
        }
        return sheetSettings;
    }

    /**
     * 获取下拉信息
     * @param workbook
     * @param sheetSettings
     * @param thisSheetIndex
     */
    private void getSelectSheetMap(Workbook workbook,SheetSettings sheetSettings,int thisSheetIndex){
        List<Name> list = (List<Name>) workbook.getAllNames();
        for (Name name: list){
            if (name.getRefersToFormula().indexOf("_" + thisSheetIndex + "_") != 0){
                sheetSettings.getSelectMap().put(name.getNameName(),new ArrayList<String>());
                sheetSettings.getSelectMap().put(name.getNameName() + "_value",new ArrayList<String>());

                Sheet textSheet = workbook.getSheet(name.getRefersToFormula().split("!")[0]);
                Sheet valueSheet = workbook.getSheet(name.getRefersToFormula().split("!")[0].replace("_text","_value"));

                String address = name.getRefersToFormula().split("!")[1];

                int rowNum = Integer.valueOf(address.split(":")[0].substring(address.split(":")[0].lastIndexOf("$") + 1));
                String[] cellAddress = address.replaceAll("['$]","").replaceAll(String.valueOf(rowNum),"").split(":");

                Row textRow = textSheet.getRow(rowNum - 1);
                Row valueRow = valueSheet.getRow(rowNum - 1);

                if (cellAddress.length > 1){
                    for (int cellIndex = CellReference.convertColStringToIndex(cellAddress[0]); cellIndex <= CellReference.convertColStringToIndex(cellAddress[1]); cellIndex++) {
                        Cell textCell = textRow.getCell(cellIndex);
                        Cell valueCell = valueRow.getCell(cellIndex);
                        if (textCell != null){
                            sheetSettings.getSelectMap().get(name.getNameName()).add(textCell.getStringCellValue());
                            sheetSettings.getSelectMap().get(name.getNameName() + "_value").add(valueCell.getStringCellValue());
                        }
                    }
                }
            }
        }
    }

    /**
     * 清理缓存
     * @param sheetSettings
     */
    private void cleanCacheData(SheetSettings[] sheetSettings){
        for (SheetSettings thisSheetSettings: sheetSettings){
            thisSheetSettings.getSelectTargetSet().clear();
            thisSheetSettings.getSelectMap().clear();
        }
    }

    private Set<String> setformulaGroupName(Set<String> group,String addName){
        if (group == null){
            group = new HashSet<String>();
        }
        group.add(addName);
        return group;
    }

    private String createFormulaByGroup(SheetSettings sheetSettings,Cell cell,FormulaType formulaType,Set<String> group){
        StringBuffer formulaStr = new StringBuffer();
        formulaStr.append(formulaType.getValue());
        formulaStr.append("(");
        int i = 0;
        for (String name: group){
            if (i++ > 0) {
                formulaStr.append(",");
            }
            formulaStr.append(sheetSettings.getCellAddressMap().get(name) + (cell.getRowIndex() + 1));
        }
        formulaStr.append(")");
        return formulaStr.toString();
    }
}
