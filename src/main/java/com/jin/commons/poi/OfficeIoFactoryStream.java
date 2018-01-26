package com.jin.commons.poi;

import com.jin.commons.poi.model.FormulaType;
import com.jin.commons.poi.model.SheetSettings;
import com.jin.commons.poi.model.TableSettings;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.*;

/**
 * @author wujinglei
 */
public class OfficeIoFactoryStream {

    private final static Logger log = LoggerFactory.getLogger(OfficeIoFactoryStream.class);

    /**
     * 导入XLSX
     *
     * @param file   the file
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
        } catch (InvalidFormatException e) {
            log.error(e.getMessage());
        } catch (FileNotFoundException e) {
            log.error(e.getMessage());
        } catch (IOException e) {
            log.error(e.getMessage());
        }
        return scanData(workbook, sheetSettingsList);
    }

    /**
     * Import xlsx office io result.
     *
     * @param inputStream the input stream
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
        } catch (InvalidFormatException e) {
            log.error(e.getMessage());
        } catch (FileNotFoundException e) {
            log.error(e.getMessage());
        } catch (IOException e) {
            log.error(e.getMessage());
        }
        return scanData(workbook, sheetSettingsList);
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

    private OfficeIoResult scanData(Workbook workbook,List<SheetSettings> sheetSettingsList){
        return null;
    }

    /**
     * 导出异常记录
     * @return
     */
    protected final OfficeIoResult exportXlsxErrorRecord(){
        return null;
    }

    /**
     * 导出模板
     * @return
     */
    protected final OfficeIoResult exportXlsxTemplate(List<SheetSettings> sheetSettingsList){
        sheetSettingsList.stream().forEach(thisSheetSettings -> {
            initSheetSettings();
        });
        return null;
    }

    /**
     * 列表
     * @return
     */
    protected final OfficeIoResult exportXlsx(List<SheetSettings> sheetSettingsList){
        sheetSettingsList.stream().forEach(thisSheetSettings -> {
            initSheetSettings();
        });
        return null;
    }


    /**
     * 创建SHEET
     */
    private void createSheet(){
        initSheetSettings();
        createTable();
        graffitiData();
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
    private void createTable(){
        initTableSettings();
        createTableTitle();
        createHeader();
    }

    /**
     * 创建表格标题
     */
    private void createTableTitle(){

    }

    /**
     * 填写数据
     */
    private void graffitiData(){

    }

    private void createHeader(){

    }

    /**
     * 初始化SHEET配置
     */
    private void initSheetSettings (){

    }

    /**
     * 初始化TABLE配置
     */
    private void initTableSettings (){

    }

    /**
     * 初始化Cell信息
     */
    private void initCellSettings (){

    }

    /**
     * 将数据 放入对象
     */
    private void setValueToObject (){

    }

    /**
     * 读取单元格数据
     * @return
     */
    private Object getCellValue(){
        return null;
    }

    /**
     * 读取对象中的数据
     */
    private String getValueFromObj(){
        return null;
    }

    /**
     * 创建表头
     * @return
     */
    private Cell createHeaderCell(){
        return null;
    }

    /**
     * 创建单元格
     * @return
     */
    private Cell createDataCell(){
        return null;
    }

    /**
     * 将数据放入单元格
     */
    private void setCellDataValue(){

    }

    /**
     * 生成样式
     * @return
     */
    private CellStyle createCellStyle(){
        return null;
    }

    /**
     * 记录异常信息
     */
    private void recordSetCellDataValueException(){

    }

    /**
     * 处理下拉列表问题
     */
    private void createHideSelectSheet(){

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
        XSSFDataValidationConstraint dvConstraint = new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.LIST,formulaString);
        CellRangeAddressList addressList = new CellRangeAddressList(rowIndex, rowIndex, xlsCellIndex, xlsCellIndex);
        DataValidation dataValidation = sheet.getDataValidationHelper().createValidation(dvConstraint, addressList);
        dataValidation.setShowErrorBox(true);
        sheet.addValidationData(dataValidation);
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

    private Set<String> setformulaGroupName(Set<String> group, String addName){
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
     * init cell Rules
     * @author: wujinglei
     * @date: 2014年7月8日 下午4:46:02
     * @Description: 判断规则
     */
    private Boolean initRule(){
        return true;
    }
}
