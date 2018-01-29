package com.jin.commons.poi;

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
import org.apache.poi.xssf.usermodel.XSSFName;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * The type Office io factory stream.
 *
 * @author wujinglei
 */
public class OfficeIoFactory {

    private final static Logger log = LoggerFactory.getLogger(OfficeIoFactory.class);

    /**
     * 导入XLSX
     *
     * @param file              the file
     * @param sheetSettingsList the sheet settings list
     * @return office io result
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:24:29
     * @Description: 导入XLSX
     */
    protected final OfficeIoResult importXlsx(File file,  List<SheetSettings> sheetSettingsList) {
        // 按文件取出工作簿
        Workbook workbook = null;
        try {
            workbook = create(new FileInputStream(file));
        } catch (InvalidFormatException | IOException e) {
            log.error(e.getMessage());
        }
        return loadWorkbook(workbook, sheetSettingsList);
    }

    /**
     * Import xlsx office io result.
     *
     * @param inputStream       the input stream
     * @param sheetSettingsList the sheet settings list
     * @return office io result
     * @author: wujinglei
     * @date: 2014年6月11日 上午10:24:29
     * @Description: 导入XLS
     */
    protected final OfficeIoResult importXlsx(InputStream inputStream, List<SheetSettings> sheetSettingsList) {
        // 按文件取出工作簿
        Workbook workbook = null;
        try {
            workbook = create(inputStream);
        } catch (InvalidFormatException | IOException e) {
            log.error(e.getMessage());
        }
        return loadWorkbook(workbook, sheetSettingsList);
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
     * 读取配置
     * @param workbook
     * @param sheetSettingsList
     * @return
     */
    private OfficeIoResult loadWorkbook(Workbook workbook,List<SheetSettings> sheetSettingsList){
        OfficeIoResult result = new OfficeIoResult(sheetSettingsList);
        result.setresultWorkbook(workbook);
        //文件异常时处理
        if (workbook == null) {
            result.addErrorRecord(new ErrorRecord("文件无法读取或读取异常", "跳过所有处理", true));
            return result;
        }

        // check selectSheet
        getSelectSheetMap(result);

        sheetSettingsList = initSheetSettings(sheetSettingsList);

        sheetSettingsList.parallelStream().forEach(thisSheetSettings -> loadSheet(result, thisSheetSettings));

//        cleanCacheData(sheets);

        result.setSheetSettings(sheetSettingsList);

        return result;
    }

    /**
     * 导出异常记录
     *
     * @return office io result
     */
    protected final OfficeIoResult exportXlsxErrorRecord(){
//        //实例化返回对象
//        OfficeIoResult result = new OfficeIoResult(sheets);
//        //循环构建sheet
//        Set<Integer> keySet = errRecordRows.keySet();
//        for (Integer index : keySet) {
//            try{
//
//                SheetSettings thisSheetSettings = checkCellSettings(result.getResultWorkbook(), sheets[index], index);
//
//                //创建sheet
//                Sheet sheet = result.getResultWorkbook().createSheet(sheets[index].getSheetName());
//
//                boolean hasSubTitle = buildHeader(result.getResultWorkbook(), sheet, thisSheetSettings);
//
//                int startRow = hasSubTitle?1:0;
//                //写入出错行记录
//                List rowList = errRecordRows.get(index);
//                int errorRowCount = rowList.size();
//                for (int errorIndex = 0; errorIndex < errorRowCount; errorIndex++) {
//                    Row row = sheet.createRow(errorIndex + 1 + startRow);
//                    if (rowList.get(errorIndex) instanceof Row) {
//                        Iterator<Cell> it = ((Row) rowList.get(errorIndex)).cellIterator();
//                        int cellIndex = 0;
//                        while (it.hasNext()) {
//                            Cell sourceCell = it.next();
//                            Cell targetCell = row.createCell(cellIndex++);
//                            targetCell.setCellType(CellType.STRING);
//                            if (CellType.NUMERIC.equals(sourceCell.getCellTypeEnum())) {
//                                targetCell.setCellValue(sourceCell.getNumericCellValue());
//                            }
//                            if (CellType.STRING.equals(sourceCell.getCellTypeEnum())) {
//                                targetCell.setCellValue(sourceCell.getRichStringCellValue());
//                            }
//                            if (CellType.FORMULA.equals(sourceCell.getCellTypeEnum())) {
//                                targetCell.setCellValue(sourceCell.getCellFormula());
//                            }
//                            if (CellType.BOOLEAN.equals(sourceCell.getCellTypeEnum())) {
//                                targetCell.setCellValue(sourceCell.getBooleanCellValue());
//                            }
//                            if (CellType.ERROR.equals(sourceCell.getCellTypeEnum())) {
//                                targetCell.setCellValue(sourceCell.getErrorCellValue());
//                            }
//                        }
//                    } else if (rowList.get(errorIndex) instanceof String[]) {
//                        String[] values = (String[]) rowList.get(errorIndex);
//                        for (int i = 0; i < values.length; i++) {
//                            Cell targetCell = row.createCell(i);
//                            targetCell.setCellType(CellType.STRING);
//                            targetCell.setCellValue(values[i]);
//                        }
//                    }
//                }
//            }catch (Exception e){
//                log.error(e.getMessage(),e);
//            }
//        }
//        return result;
        return null;
    }

    /**
     * 导出模板
     *
     * @param sheetSettingsList the sheet settings list
     * @return office io result
     */
    protected final OfficeIoResult exportXlsxTemplate(List<SheetSettings> sheetSettingsList){

        // 实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheetSettingsList);

        sheetSettingsList = initSheetSettings(sheetSettingsList);

        // 处理下拉列表
        createHideSelectSheet(result.getResultWorkbook(),sheetSettingsList);

        sheetSettingsList.stream().forEach(thisSheetSettings -> createSheet(result, thisSheetSettings,true));

        return result;
    }

    /**
     * 列表
     *
     * @param sheetSettingsList the sheet settings list
     * @return office io result
     */
    protected final OfficeIoResult exportXlsx(List<SheetSettings> sheetSettingsList){

        // 实例化返回对象
        OfficeIoResult result = new OfficeIoResult(sheetSettingsList);

        sheetSettingsList = initSheetSettings(sheetSettingsList);

        // 处理下拉列表
        createHideSelectSheet(result.getResultWorkbook(),sheetSettingsList);

        sheetSettingsList.parallelStream().forEach(thisSheetSettings -> createSheet(result, thisSheetSettings,false));

        return result;
    }


    /**
     * 创建SHEET
     */
    private void createSheet(OfficeIoResult result, SheetSettings sheetSettings,Boolean isDemo){
        // 创建sheet
        Sheet sheet = result.getResultWorkbook().createSheet(sheetSettings.getSheetName());
        /*
         * 1.表格开始行
         */
        int[] args = {0};

        // TODO 获取所有TABLE的开始行

        sheetSettings.getTableSettingsList().stream().forEach(thisTableSettings -> {
            createTable(sheet,thisTableSettings,args[0]++);
            if (isDemo){
                // 导入DEMO数据
                setDemoData(sheet,sheetSettings,thisTableSettings);
            }else {
                graffitiData(result,sheet,thisTableSettings,sheetSettings);
            }
        });
    }

    /**
     * 读取SHEET页
     * @param result
     * @param sheetSettings
     */
    private void loadSheet(OfficeIoResult result,SheetSettings sheetSettings){
        sheetSettings.getTableSettingsList().stream().forEach(tableSettings -> loadTableData(
                result,
                result.getResultWorkbook().getSheet(sheetSettings.getSheetName()),
                sheetSettings,
                tableSettings
        ));
    }

    /**
     * 填写DEMO模板数据
     * @param sheet
     * @param sheetSettings
     * @param tableSettings
     */
    private void setDemoData(Sheet sheet,SheetSettings sheetSettings, TableSettings tableSettings) {
        int[] args = {tableSettings.getSkipRows()};
        List<CellSettings> cellArrays = getAllCell(tableSettings.getCellSettingsList());
        IntStream
                .range(0,5)
                .forEach(index -> {
                    Row row = sheet.createRow(index + args[0]);
                    cellArrays.parallelStream().forEach(cellSettings -> {
                        //构建一个CELL
                        Cell cell = createDataCell(row, cellSettings);
                        try {
                            setCellDataValue(sheet, cell,sheetSettings,tableSettings, cellSettings, null);
                        } catch (Exception e) {
                            log.warn(e.getMessage());
                        }
                    });
                });
    }

    /**
     * 读取SHEET配置
     */
    private void loadSheetSettings(){

    }

    /**
     * 读取TABLE配置
     */
    private void loadTablesSettings(){

    }

    /**
     * 读取CELL配置
     */
    private void loadCellSettings(){

    }

    /**
     * 创建表格
     */
    private void createTable(Sheet sheet, TableSettings tableSettings,Integer startRowNum){
        // 构建标题
        if (!StringUtils.isBlank(tableSettings.getTitle())){
            createTableTitle(sheet,tableSettings);
        }

        // 构建表头
        createHeader(sheet,tableSettings,sheet.createRow(startRowNum));
    }

    /**
     * 创建表格标题
     */
    private void createTableTitle(Sheet sheet, TableSettings tableSettings){
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(tableSettings.getTitle());
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();

        style.setFillForegroundColor(tableSettings.getTitleStyle().getTitleForegroundColor());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderRight(tableSettings.getTitleStyle().getTitleBorder()[0]);
        style.setBorderTop(tableSettings.getTitleStyle().getTitleBorder()[1]);
        style.setBorderLeft(tableSettings.getTitleStyle().getTitleBorder()[2]);
        style.setBorderBottom(tableSettings.getTitleStyle().getTitleBorder()[3]);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        font.setFontName(tableSettings.getTitleStyle().getTitleFont());
        font.setColor(tableSettings.getTitleStyle().getTitleFontColor());
        font.setFontHeightInPoints(tableSettings.getTitleStyle().getTitleSize());
        style.setFont(font);

        titleCell.setCellStyle(style);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, tableSettings.getCellCount() - 1);
        sheet.addMergedRegion(region);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
    }

    /**
     * 读取Table Data
     *
     * @param result
     * @param sheet
     * @param tableSettings
     * @param sheetSettings
     */
    private void loadTableData(OfficeIoResult result, Sheet sheet,SheetSettings sheetSettings,TableSettings tableSettings){

        List tableDataList = new ArrayList();

        /*
         * 1. rowIndex
         */
        int[] args = {tableSettings.getSkipRows()};
        while (!tableIsClosed(sheet,args[0])){
            Row activeRow = sheet.getRow(args[0]++);
            if (activeRow != null) {
                // 第一行的各列放在一个MAP中
                final Object[] resultObj = {new Object()};
                try {
                    if (tableSettings.getDataClazzType() != null){
                        resultObj[0] = tableSettings.getDataClazzType().newInstance();
                    }else {
                        resultObj[0] = new HashMap();
                    }
                } catch (InstantiationException | IllegalAccessException e) {
                    log.error(e.getMessage());
                    resultObj[0] = new HashMap();
                }
                // 循环每一列按列所给的参数进行处理
                Map<String,String> selectTargetValueMap = new HashMap();
                List<CellSettings> allCellSettings = getAllCell(tableSettings.getCellSettingsList());
                allCellSettings.parallelStream().forEach(cellSettings -> {
                    Cell cell = activeRow.getCell(cellSettings.getCellSeq());
                    if (cell != null) {
                        Object obj = null;
                        if (!cellSettings.isSkip()){
                            try {
                                obj = getCellValue(result,cell, sheetSettings,tableSettings, cellSettings,selectTargetValueMap);
                            } catch (XSSFCellTypeException e) {
                                // TODO 需重新处理异常内容
                                // recordSetCellDataValueException(result,activeRow,thisSheetSettings, cell.getAddress().formatAsString(),cellSettings,e);
//                                continue rowLoop;
                            }
                            //判断规则
                            if (!checkRule(cellSettings, cell, obj, result, sheetSettings.getSheetSeq(), activeRow)) {
//                                continue rowLoop;
                            }
                            setValueToObject(resultObj[0], cellSettings, obj);
                        }
                    }
                });

                //将前当行所对应的MAP放入List中
                tableDataList.add(resultObj);
            } else {
                result.addWrongRecord(new WrongRecord(sheetSettings.getSheetSeq(), args[0], "导入的文件中空行数据", "跳过行处理", false));
                continue;
            }
        }

//        // 循环每一行
//        rowLoop:
//        for (int row = 0; row <= rowsNum; row++) {
//            //判断是否是在skipRow之内
//            if (row < thisSheetSettings.getSkipRows()) {
//                continue;
//            }
//            // 取的当前行
//            Row activeRow = sheet.getRow(row);
//            // 判断当前行记录是否有有效
//
//        }
        //将成功条数放入result中
//        result.getResultTotal()[sheetIndex] = successCount;
        //将处理后的sheet的数据放入返回对象中
        result.addSheetList(tableDataList);

        if (result.getErrors().size() > 0){
            result.setCompleted(false);
        }
    }

    /**
     * 填写数据
     * @param result
     * @param sheet
     * @param tableSettings
     * @param sheetSettings
     */
    private void graffitiData(OfficeIoResult result, Sheet sheet,TableSettings tableSettings,SheetSettings sheetSettings){
//        //取出当前sheet所要导出的数据
//        List dataList = thisSheetSettings.getExportData();
//        CellSettings[] cells = thisSheetSettings.getCellSettings();
//
//        //循环新增每一条数据
//        long successCount = 0;
//        int startRowIndex = 1;
//        if (!StringUtils.isBlank(thisSheetSettings.getTitle())){
//            startRowIndex += 1;
//        }
//        if (hasSubTitle) {
//            startRowIndex += 1;
//        }
//
//        if (dataList != null && dataList.size() > 0) {
//            rowLoop:
//            for (int dataIndex = 0; dataIndex < dataList.size(); dataIndex++) {
//                //取出当前行的数据对象
//                Object bean = dataList.get(dataIndex);
//                //新增行
//                Row row = sheet.createRow(dataIndex + startRowIndex);
//                //循环列配置为第一列赋值
//                for (int cellIndex = 0, xlsCellIndex = 0; cellIndex < cells.length; cellIndex++) {
//                    CellSettings thisCellSettings = cells[cellIndex];
//
//                    if (cells[cellIndex].getSubCells() == null) {
//                        //构建一个CELL
//                        Cell cell = createDataCell(workbook, row, xlsCellIndex, thisCellSettings);
//                        //写入内容
//                        try {
//                            setCellDataValue(sheet, cell,thisSheetSettings, thisCellSettings, bean);
//                        } catch (Exception e) {
//                            recordSetCellDataValueException(result, row, thisSheetSettings, cell.getAddress().formatAsString(), thisCellSettings, e);
//                            continue rowLoop;
//                        }
//                        xlsCellIndex++;
//                    } else {
//                        for (int subIndex = 0; subIndex < cells[cellIndex].getSubCells().length; subIndex++) {
//                            CellSettings thisSubCellSettings = cells[cellIndex].getSubCells()[subIndex];
//                            //构建一个CELL
//                            Cell cell = createDataCell(workbook, row, xlsCellIndex, thisSubCellSettings);
//                            //写入内容
//                            try {
//                                setCellDataValue(sheet, cell, thisSheetSettings,thisSubCellSettings, bean);
//                            } catch (Exception e) {
//                                recordSetCellDataValueException(result, row, thisSheetSettings, cell.getAddress().formatAsString(), thisCellSettings, e);
//                                continue rowLoop;
//                            }
//                            xlsCellIndex++;
//                        }
//                    }
//                }
//                //记录成功结果
//                successCount++;
//            }
//        }
    }

    private void createHeader(Sheet sheet, TableSettings tableSettings, Row startRow){
        boolean hasSubTitle;
        hasSubTitle = tableSettings.getCellSettingsList().stream().anyMatch(cellSettings -> cellSettings.getSubCells() != null);
        if (hasSubTitle){
            // 有合并表头
            tableSettings.getCellSettingsList().stream().forEach(thisCellSettings -> {
                Cell cell = createHeaderCell(startRow, thisCellSettings);

                cell.setCellValue(thisCellSettings.getColName());

                if (thisCellSettings.getSubCells() != null) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow.getRowNum(), startRow.getRowNum(), thisCellSettings.getCellSeq(), thisCellSettings.getCellSeq() + thisCellSettings.getSubCells().size() - 1));
                }
            });

            Row subRow = sheet.createRow(startRow.getRowNum() + 1);
            tableSettings.getCellSettingsList().parallelStream().forEach(thisCellSettings -> {
                if (thisCellSettings.getSubCells() != null) {
                    thisCellSettings.getSubCells().parallelStream().forEach(subCellSettings -> {
                        Cell subTitleCell = createHeaderCell(subRow, subCellSettings);
                        subTitleCell.setCellValue(subCellSettings.getColName());
                    });
                } else {
                    CellRangeAddress region = new CellRangeAddress(startRow.getRowNum(), startRow.getRowNum() + 1, thisCellSettings.getCellSeq(), thisCellSettings.getCellSeq());
                    sheet.addMergedRegion(region);
                    // 处理合并单元格的边框问题
                    RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
                }
            });
        }else {
            // 无合并表头
            tableSettings.getCellSettingsList().parallelStream().forEach(cellSettings -> {
                Cell subTitleCell = createHeaderCell(startRow, cellSettings);
                subTitleCell.setCellValue(cellSettings.getColName());
            });
        }
    }

    /**
     * 初始化SHEET配置
     */
    private List<SheetSettings> initSheetSettings (List<SheetSettings> sheetSettingsList){
        /*
         *  0.sheetSeq
         *  1.tableSeq
         */
        final int[] args = {0,0};
        sheetSettingsList = sheetSettingsList.stream().peek(thisSheetSettings -> {
            // setSheetSeq
            thisSheetSettings.setSheetSeq(args[0]++);
            // setTableSeq
            List<TableSettings> newTablesList = thisSheetSettings.getTableSettingsList().stream().peek(thisTableSettings -> {
                thisTableSettings.setTableSeq(args[1]++);
                // setCellSeq
                final int[] cellIndex = {0};
                List<CellSettings> newCellSettingList = thisTableSettings.getCellSettingsList().stream().peek(thisCellSettings -> {
                    if (thisCellSettings.getSubCells() != null) {
                        thisCellSettings.setCellSeq(cellIndex[0]);
                        List<CellSettings> subCellSettingsList =
                                thisCellSettings.getSubCells()
                                        .stream()
                                        .peek(subCellSettings -> {
                                            thisTableSettings.getCellAddressMap().put(subCellSettings.getKey(),CellReference.convertNumToColString(cellIndex[0]));
                                            subCellSettings.setCellSeq(cellIndex[0]++);
                                        })
                                        .collect(Collectors.toList());
                        thisCellSettings.setSubCells(subCellSettingsList);
                    } else {
                        thisTableSettings.getCellAddressMap().put(thisCellSettings.getKey(),CellReference.convertNumToColString(cellIndex[0]));
                        thisCellSettings.setCellSeq(cellIndex[0]++);
                    }
                }).collect(Collectors.toList());

                thisTableSettings.setCellCount(cellIndex[0]);
                thisTableSettings.setCellSettings(newCellSettingList);
            }).collect(Collectors.toList());
            thisSheetSettings.getTableSettingsList(newTablesList);
        }).collect(Collectors.toList());

        sheetSettingsList =
                sheetSettingsList.parallelStream()
                        .peek(thisSheetSettings ->
                                initTableSettings(thisSheetSettings.getTableSettingsList())
                        )
                        .collect(Collectors.toList());
        return sheetSettingsList;
    }

    /**
     * 初始化TABLE配置
     */
    private List<TableSettings> initTableSettings (List<TableSettings> tableSettingsList) {
        tableSettingsList = tableSettingsList.parallelStream().map(thisTableSettings -> {
            // checkSkipRow
            if (thisTableSettings.getSkipRows() == null) {
                thisTableSettings.setSkipRows(1);
            }
            if (!StringUtils.isBlank(thisTableSettings.getTitle())){
                thisTableSettings.setSkipRows(thisTableSettings.getSkipRows() + 1);
            }
            if (thisTableSettings.getCellSettingsList()
                    .parallelStream()
                    .anyMatch(cellSettings -> cellSettings.getSubCells() != null)){
                thisTableSettings.setSkipRows(thisTableSettings.getSkipRows() + 1);
            }

            // 初始化Cell信息
            return initCellSettings(thisTableSettings);
        }).collect(Collectors.toList());
        return tableSettingsList;
    }

    /**
     * 初始化Cell信息
     */
    private TableSettings initCellSettings (TableSettings tableSettings) {
        Map<String,Set<String>> formulaMap = new HashMap<String, Set<String>>();
        List<CellSettings> cellArrays = getAllCell(tableSettings.getCellSettingsList());

        // 处理联动下拉的Target问题
        cellArrays.stream().forEach(thisCellSettings -> {
            if (thisCellSettings.getCellDataType() != CellDataType.FORMULA){
                if (thisCellSettings.getFormulaGroupNames() != null){
                    Arrays.stream(thisCellSettings.getFormulaGroupNames())
                            .forEach(
                                    groupName ->
                                            formulaMap.put(groupName, setFormulaGroupName(formulaMap.get(groupName),thisCellSettings.getKey()))
                            );
                }
            }

            if (thisCellSettings.getSelectCascadeFlag()){
                tableSettings.getSelectTargetSet().add(thisCellSettings.getSelectTargetKey());
            }
        });

        // 配置公式
        cellArrays.parallelStream().forEach(thisCellSettings -> {
            if (thisCellSettings.getCellDataType() == CellDataType.FORMULA){
                thisCellSettings.getFormulaSettings().setGroupName(formulaMap.get(thisCellSettings.getFormulaGroupNames()[0]));
            }
        });

        // 读取CLASS配置
        if (tableSettings.getDataClazzType() != null){
            tableSettings.setCellSettings(
                tableSettings.getCellSettingsList().stream().peek(thisCellSettings -> {
                    if (thisCellSettings.getSubCells() == null){
                        if (thisCellSettings.getCellDataType() != CellDataType.FORMULA){
                            if (thisCellSettings.getCellClass() == null){
                                thisCellSettings.setCellClass(FieldUtils.getDeclaredFieldType(tableSettings.getDataClazzType(),thisCellSettings.getKey()));
                                thisCellSettings.setCellDataType(FieldUtils.getCellDataType(thisCellSettings.getCellClass()));
                            }
                        }
                    }else {
                        thisCellSettings.setSubCells(
                            thisCellSettings.getSubCells().stream().peek(subCellSettings -> {
                                if (subCellSettings.getCellDataType() != CellDataType.FORMULA){
                                    subCellSettings.setCellClass(FieldUtils.getDeclaredFieldType(tableSettings.getDataClazzType(),subCellSettings.getKey()));
                                    subCellSettings.setCellDataType(FieldUtils.getCellDataType(subCellSettings.getCellClass()));
                                }
                            }).collect(Collectors.toList())
                        );
                    }
                }).collect(Collectors.toList())
            );
        }

        return tableSettings;
    }

    /**
     * 获取所有CELL对象
     * @param cellSettingsList
     * @return
     */
    private List<CellSettings> getAllCell(List<CellSettings> cellSettingsList){
        List<CellSettings> cellArrays = new ArrayList<CellSettings>();
        cellSettingsList.parallelStream()
                .filter(thisCellSetting -> thisCellSetting.getSubCells() == null)
                .forEach(cellArrays::add);
        cellSettingsList.parallelStream()
                .filter(thisCellSetting -> thisCellSetting.getSubCells() != null)
                .forEach(thisCellSetting -> cellArrays.addAll(thisCellSetting.getSubCells()));
        return cellArrays;
    }

    /**
     * 将数据 放入对象
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
     * @return
     */
    private Object getCellValue(OfficeIoResult result,
                                Cell cell,
                                SheetSettings sheetSettings,
                                TableSettings tableSettings,
                                CellSettings cellSettings,
                                Map<String,String> selectTargetValueMap) throws XSSFCellTypeException {
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


        if (tableSettings.getSelectTargetSet().size() > 0){
            if (tableSettings.getSelectTargetSet().contains(cellSettings.getKey())){
                selectTargetValueMap.put(cellSettings.getKey(),cellValue);
            }
        }

        // 处理下拉选择问题
        if (cellSettings.getSelect()){
            if (!cellSettings.getSelectCascadeFlag()){
                String formulaString = cellSettings.getKey() + "_TEXT";
                List<String> mapList = (List<String>) result.getSelectMap().get(DigestUtils.digestFormulaName(formulaString));
                int matchIndex = mapList.indexOf(cellValue);
                if (matchIndex != -1){
                    cellValue = ((List<String>)result.getSelectMap().get(DigestUtils.digestFormulaName(formulaString) + "_value")).get(matchIndex);
                }else{
                    // TODO warn
                }
            }else {
                String formulaString = cellSettings.getKey() + "_" + selectTargetValueMap.get(cellSettings.getSelectTargetKey()) + "_TEXT";
                List<String> mapList = (List<String>)result.getSelectMap().get(DigestUtils.digestFormulaName(formulaString));
                int matchIndex = mapList.indexOf(cellValue);
                if (matchIndex != -1){
                    cellValue = ((List<String>)result.getSelectMap().get(DigestUtils.digestFormulaName(formulaString) + "_value")).get(matchIndex);
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
                        FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
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
     */
    private String getValueFromObj(CellSettings cellSettings, Object bean) {
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
     * 创建表头
     * @return
     */
    private Cell createHeaderCell(Row row, CellSettings cellSettings){
        // 构建一个CELL
        Cell cell = row.createCell(cellSettings.getCellSeq());
        // 设置CELL为文本格式
        cell.setCellType(CellType.STRING);

        cell.setCellStyle(createCellStyle(row.getSheet().getWorkbook(), cellSettings, true));

        return cell;
    }

    /**
     * 创建单元格
     * @return
     */
    private Cell createDataCell(Row row,CellSettings cellSettings){
        // 构建一个CELL
        Cell cell = row.createCell(cellSettings.getCellSeq());
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
        cell.setCellStyle(createCellStyle(row.getSheet().getWorkbook(), cellSettings, false));
        return cell;
    }

    /**
     * 将数据放入单元格
     */
    private void setCellDataValue(Sheet sheet,Cell cell,SheetSettings sheetSettings,TableSettings tableSettings,CellSettings cellSettings, Object dataBean){
        //写入内容
        if (cellSettings.getHasStaticValue()) {
            cell.setCellValue(cellSettings.getStaticValue());
        }
        if (cellSettings.getSelect()) {
            StringBuilder formulaString = new StringBuilder();
            if (cellSettings.getSelectCascadeFlag()){
                String addressFlag = tableSettings.getCellAddressMap().get(cellSettings.getSelectTargetKey());
                // =INDIRECT(VLOOKUP(A1,Sheet2!A:B,2,0))
                formulaString.append("INDIRECT(VLOOKUP(");
                formulaString.append("CONCATENATE(\"");
                formulaString.append(cellSettings.getKey());
                formulaString.append("_\",");
                formulaString.append(addressFlag);
                formulaString.append(cell.getAddress().getRow() + 1);
                formulaString.append("),select_text");
                formulaString.append("!A:B,2,0))");
            }else {
                formulaString.append(DigestUtils.digestFormulaName(sheetSettings.getSheetSeq() + "_" + tableSettings.getTableSeq() + "_" + cellSettings.getKey() + "_TEXT"));
            }
            setSelectDataValidation(sheet,formulaString.toString(),cell.getRowIndex(),cell.getColumnIndex());
        }

        if (dataBean != null) {
            if (cellSettings.getCellDataType() != CellDataType.FORMULA){
                String reVal = getValueFromObj(cellSettings, dataBean);
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
                    String formulaStr = createFormulaByGroup(tableSettings,cell,formulaSettings.getFormulaType(),formulaSettings.getGroupName());
                    cell.setCellFormula(formulaStr);
                }

            }
        }
    }

    /**
     * 生成样式
     * @return
     */
    private CellStyle createCellStyle(Workbook workbook, CellSettings cellSettings, boolean isTitle){
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
     * 记录异常信息
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
     * @param sheetSettingsList
     */
    private void createHideSelectSheet(Workbook workbook, List<SheetSettings> sheetSettingsList) {
        Sheet selectTextSheet = workbook.createSheet("select_text");
        Sheet selectValueSheet = workbook.createSheet("select_value");

        /*
         * 1. selectRowIndex
         */
        int[] args = {0};
        // 先处理没有联动的下拉
        Map<String,String[]> textMapping = new HashMap<String, String[]>();
        Map<String,String[]> valueMapping = new HashMap<String, String[]>();

        sheetSettingsList.stream().forEach(sheetSettings -> sheetSettings.getTableSettingsList().stream().forEach(tableSettings -> {
            List<CellSettings> cellArrays = getAllCell(tableSettings.getCellSettingsList());
            // 先处理没有联动的下拉
            cellArrays.stream()
                    .filter(cellSettings -> cellSettings.getSelect() && !cellSettings.getSelectCascadeFlag())
                    .forEach(cellSettings -> {
                        int thisSelectRowIndex = args[0]++;
                        setSelectRow(selectTextSheet, selectValueSheet, thisSelectRowIndex, cellSettings.getSelectValueList(), cellSettings.getSelectTextList(),false);
                        createSelectNameList(selectTextSheet.getSheetName(), workbook, DigestUtils.digestFormulaName(sheetSettings.getSheetSeq() + "_" + tableSettings.getTableSeq() + "_" + cellSettings.getKey() + "_TEXT"), thisSelectRowIndex, cellSettings.getSelectTextList().length, cellSettings.getSelectCascadeFlag());
                        textMapping.put(cellSettings.getKey(),cellSettings.getSelectTextList());
                        valueMapping.put(cellSettings.getKey(),cellSettings.getSelectValueList());
                    });

            // 处理有联动的下拉
            cellArrays.stream()
                    .filter(cellSettings -> cellSettings.getSelect() && cellSettings.getSelectCascadeFlag())
                    .forEach(cellSettings -> {
                        if (cellSettings.getSelectSourceList() != null && cellSettings.getSelectSourceList().size() > 0){
                            String[] targetTextArray = textMapping.get(cellSettings.getSelectTargetKey());
                            String[] targetValueArray = valueMapping.get(cellSettings.getSelectTargetKey());
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

                                    cellSettings.getSelectSourceList().stream().forEach(obj -> {
                                        int matchIndex = ArrayUtils.indexOf(targetValueArray, String.valueOf(BeanUtils.invokeGetter(obj, cellSettings.getBingKey())));
                                        if (matchIndex >= 0) {
                                            addTextSelectMap.get(targetTextArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, cellSettings.getCellSelectSettings().getKey())));
                                            addValueSelectMap.get(targetValueArray[matchIndex]).add(String.valueOf(BeanUtils.invokeGetter(obj, cellSettings.getCellSelectSettings().getName())));
                                        }
                                    });

                                    cellSettings.getSelectSourceList().stream().forEach(obj -> {
                                        int matchIndex = ArrayUtils.indexOf(targetValueArray, String.valueOf(BeanUtils.invokeGetter(obj, cellSettings.getBingKey())));
                                        if (matchIndex >= 0) {
                                            String formulaStr = DigestUtils.digestFormulaName(cellSettings.getKey() + "_" + System.nanoTime() + "_" + "_TEXT");
                                            String[] addTextArray = new String[addTextSelectMap.get(targetTextArray[matchIndex]).size()];
                                            addTextSelectMap.get(targetTextArray[matchIndex]).toArray(addTextArray);
                                            String[] addValueArray = new String[addValueSelectMap.get(targetValueArray[matchIndex]).size()];
                                            addValueSelectMap.get(targetValueArray[matchIndex]).toArray(addValueArray);
                                            addTextArray[0] = cellSettings.getKey() + "_" + addTextArray[0];
                                            addTextArray = ArrayUtils.insert(1,addTextArray,formulaStr);
                                            addValueArray[0] = cellSettings.getKey() + "_" + addValueArray[0];
                                            addValueArray = ArrayUtils.insert(1,addValueArray,formulaStr);
                                            setSelectRow(selectTextSheet, selectValueSheet, args[0], addTextArray, addValueArray,true);
                                            createSelectNameList(selectTextSheet.getSheetName(), workbook, formulaStr, args[0]++, addTextSelectMap.get(targetTextArray[matchIndex]).size(), cellSettings.getSelectCascadeFlag());
                                        }
                                    });
                                }
                            }
                        }
                    });
        }));

        workbook.setSheetHidden(workbook.getSheetIndex("select_text"), false);
        workbook.setSheetHidden(workbook.getSheetIndex("select_value"), false);
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
     * @param selectTextList
     * @param selectValueList
     * @param cascadeFlag
     */
    private void setSelectRow(Sheet selectTextSheet, Sheet selectValueSheet, int selectRowIndex, String[] selectValueList, String[] selectTextList, boolean cascadeFlag) {
        createSelectRow(selectTextSheet.createRow(selectRowIndex), selectTextList, cascadeFlag);
        createSelectRow(selectValueSheet.createRow(selectRowIndex), selectValueList, cascadeFlag);
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
        XSSFDataValidationConstraint dvConstraint = new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.LIST,formulaString);
        CellRangeAddressList addressList = new CellRangeAddressList(rowIndex, rowIndex, xlsCellIndex, xlsCellIndex);
        DataValidation dataValidation = sheet.getDataValidationHelper().createValidation(dvConstraint, addressList);
        dataValidation.setShowErrorBox(true);
        sheet.addValidationData(dataValidation);
    }

    /**
     * 获取下拉信息
     * @param officeIoResult
     */
    private void getSelectSheetMap(OfficeIoResult officeIoResult){
        List<XSSFName> list = officeIoResult.getResultWorkbook().getAllNames();
        list.parallelStream().forEach(name -> {
            officeIoResult.getSelectMap().put(name.getNameName(),new ArrayList<String>());
            officeIoResult.getSelectMap().put(name.getNameName() + "_value",new ArrayList<String>());

            Sheet textSheet = officeIoResult.getResultWorkbook().getSheet(name.getRefersToFormula().split("!")[0]);
            Sheet valueSheet = officeIoResult.getResultWorkbook().getSheet(name.getRefersToFormula().split("!")[0].replace("_text","_value"));

            String address = name.getRefersToFormula().split("!")[1];

            int rowNum = Integer.valueOf(address.split(":")[0].substring(address.split(":")[0].lastIndexOf("$") + 1));
            String[] cellAddress = address.replaceAll("['$]","").replaceAll(String.valueOf(rowNum),"").split(":");

            Row textRow = textSheet.getRow(rowNum - 1);
            Row valueRow = valueSheet.getRow(rowNum - 1);

            if (cellAddress.length > 1){
                IntStream
                        .rangeClosed(CellReference.convertColStringToIndex(cellAddress[0]),CellReference.convertColStringToIndex(cellAddress[1]))
                        .forEach(cellIndex -> {
                            Cell textCell = textRow.getCell(cellIndex);
                            Cell valueCell = valueRow.getCell(cellIndex);
                            if (textCell != null){
                                ((List)officeIoResult.getSelectMap().get(name.getNameName())).add(textCell.getStringCellValue());
                                ((List)officeIoResult.getSelectMap().get(name.getNameName() + "_value")).add(valueCell.getStringCellValue());
                            }
                        });
            }
        });
    }

    /**
     * 设置公式组名称
     * @param group
     * @param addName
     * @return
     */
    private Set<String> setFormulaGroupName(Set<String> group, String addName){
        if (group == null){
            group = new HashSet<String>();
        }
        group.add(addName);
        return group;
    }

    private String createFormulaByGroup(TableSettings tableSettings, Cell cell, FormulaType formulaType, Set<String> group){
        StringBuffer formulaStr = new StringBuffer();
        formulaStr.append(formulaType.getValue());
        formulaStr.append("(");
        int i = 0;
        for (String name: group){
            if (i++ > 0) {
                formulaStr.append(",");
            }
            formulaStr.append(tableSettings.getCellAddressMap().get(name) + (cell.getRowIndex() + 1));
        }
        formulaStr.append(")");
        return formulaStr.toString();
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

    private boolean tableIsClosed(Sheet sheet,Integer rowNum){
        Row row = sheet.getRow(rowNum);
        if (row != null) {
            if (row.getRowNum() == row.getSheet().getLastRowNum()){
                return true;
            }else {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getStringCellValue().contains("[-----]")){
                    return true;
                }else {
                    return false;
                }
            }
        }else {
            return true;
        }
    }
}
