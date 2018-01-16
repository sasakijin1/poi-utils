package demo;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.FileOutputStream;
import java.util.HashMap;


public class ExcelSelectUtils {

    private static String EXCEL_HIDE_SHEET_NAME = "excelhidesheetname";
    private static String HIDE_SHEET_NAME_SEX = "sexList";
    private static String HIDE_SHEET_NAME_PROVINCE = "provinceList";

    private HashMap map = new HashMap();
    //设置下拉列表的内容
    private static String[] sexList = {"男","女"};
    private static String[] provinceList = {"浙江","山东","江西","江苏","四川"};
    private static String[] zjProvinceList = {"浙江","杭州","宁波","温州"};
    private static String[] sdProvinceList = {"山东","济南","青岛","烟台"};
    private static String[] jxProvinceList = {"江西","南昌","新余","鹰潭","抚州"};
    private static String[] jsProvinceList = {"江苏","南京","苏州","无锡"};
    private static String[] scProvinceList = {"四川","成都","绵阳","自贡"};

    public static void main(String[] args) {
        //使用事例
        Workbook wb = new HSSFWorkbook();
        createExcelMo(wb);
        creatExcelHidePage(wb);
        setDataValidation(wb);
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream("d://excel_template.xls");
            wb.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static void  createExcelMo(Workbook wb){
        Sheet sheet = wb.createSheet("用户分类添加批导");
        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("手机号码");
        cell.setCellStyle(getTitleStyle(wb));
        cell = row.createCell(1);
        cell.setCellValue("所属父类");
        cell.setCellStyle(getTitleStyle(wb));
        cell = row.createCell(2);
        cell.setCellValue("所属子类");
        cell.setCellStyle(getTitleStyle(wb));
        cell = row.createCell(3);
    }
    /**
     * 设置模板文件的横向表头单元格的样式
     * @param wb
     * @return
     */
    private static CellStyle getTitleStyle(Workbook wb){
        CellStyle style = wb.createCellStyle();
        //对齐方式设置
        //边框颜色和宽度设置
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        //粗体字设置
        Font font = wb.createFont();
        style.setFont(font);
        return style;
    }
    /**
     * 设置模板文件的横向表头单元格的样式
     * @param wb
     * @return
     */
    public static void creatExcelHidePage(Workbook workbook){
        Sheet hideInfoSheet = workbook.createSheet(EXCEL_HIDE_SHEET_NAME);//隐藏一些信息
        //在隐藏页设置选择信息
        //第一行设置性别信息
        Row sexRow = hideInfoSheet.createRow(0);
        creatRow(sexRow, sexList);
        //第二行设置省份名称列表
        Row provinceNameRow = hideInfoSheet.createRow(1);
        creatRow(provinceNameRow, provinceList);
        //以下行设置城市名称列表
        Row cityNameRow = hideInfoSheet.createRow(2);
        creatRow(cityNameRow, zjProvinceList);

        cityNameRow = hideInfoSheet.createRow(3);
        creatRow(cityNameRow, sdProvinceList);

        cityNameRow = hideInfoSheet.createRow(4);
        creatRow(cityNameRow, jxProvinceList);

        cityNameRow = hideInfoSheet.createRow(5);
        creatRow(cityNameRow, jsProvinceList);

        cityNameRow = hideInfoSheet.createRow(6);
        creatRow(cityNameRow, scProvinceList);
        //名称管理

        //第一行设置性别信息
        creatExcelNameList(workbook, HIDE_SHEET_NAME_SEX, 1, sexList.length, false);
        //第二行设置省份名称列表
        creatExcelNameList(workbook, HIDE_SHEET_NAME_PROVINCE, 2, provinceList.length, false);
        //以后动态大小设置省份对应的城市列表
        creatExcelNameList(workbook, provinceList[0], 3, zjProvinceList.length, true);
        creatExcelNameList(workbook, provinceList[1], 4, sdProvinceList.length, true);
        creatExcelNameList(workbook, provinceList[2], 5, jxProvinceList.length, true);
        creatExcelNameList(workbook, provinceList[3], 6, jsProvinceList.length, true);
        creatExcelNameList(workbook, provinceList[4], 7, scProvinceList.length, true);
        //设置隐藏页标志
        workbook.setSheetHidden(workbook.getSheetIndex(EXCEL_HIDE_SHEET_NAME), true);
    }

    /**
     * 创建一个名称
     * @param workbook
     */
    private static void creatExcelNameList(Workbook workbook,String nameCode,int order,int size,boolean cascadeFlag){
        Name name;
        name = workbook.createName();
        name.setNameName(nameCode);
        name.setRefersToFormula(EXCEL_HIDE_SHEET_NAME+"!"+creatExcelNameList(order,size,cascadeFlag));
    }

    /**
     * 名称数据行列计算表达式
     * @param workbook
     */
    private static String creatExcelNameList(int order,int size,boolean cascadeFlag){
        char start = 'A';
        if(cascadeFlag){
            start = 'B';
            if(size<=25){
                char end = (char)(start+size-1);
                return "$"+start+"$"+order+":$"+end+"$"+order;
            }else{
                char endPrefix = 'A';
                char endSuffix = 'A';
                if((size-25)/26==0||size==51){//26-51之间，包括边界（仅两次字母表计算）
                    if((size-25)%26==0){//边界值
                        endSuffix = (char)('A'+25);
                    }else{
                        endSuffix = (char)('A'+(size-25)%26-1);
                    }
                }else{//51以上
                    if((size-25)%26==0){
                        endSuffix = (char)('A'+25);
                        endPrefix = (char)(endPrefix + (size-25)/26 - 1);
                    }else{
                        endSuffix = (char)('A'+(size-25)%26-1);
                        endPrefix = (char)(endPrefix + (size-25)/26);
                    }
                }
                return "$"+start+"$"+order+":$"+endPrefix+endSuffix+"$"+order;
            }
        }else{
            if(size<=26){
                char end = (char)(start+size-1);
                return "$"+start+"$"+order+":$"+end+"$"+order;
            }else{
                char endPrefix = 'A';
                char endSuffix = 'A';
                if(size%26==0){
                    endSuffix = (char)('A'+25);
                    if(size>52&&size/26>0){
                        endPrefix = (char)(endPrefix + size/26-2);
                    }
                }else{
                    endSuffix = (char)('A'+size%26-1);
                    if(size>52&&size/26>0){
                        endPrefix = (char)(endPrefix + size/26-1);
                    }
                }
                return "$"+start+"$"+order+":$"+endPrefix+endSuffix+"$"+order;
            }
        }
    }

    /**
     * 创建一列数据
     * @param currentRow
     * @param textList
     */
    private static void creatRow(Row currentRow,String[] textList){
        if(textList!=null&&textList.length>0){
            int i = 0;
            for(String cellValue : textList){
                Cell userNameLableCell = currentRow.createCell(i++);
                userNameLableCell.setCellValue(cellValue);
            }
        }
    }/**
     * 添加数据验证选项
     * @param sheet
     */
    public static void setDataValidation(Workbook wb){
        int sheetIndex = wb.getNumberOfSheets();
        if(sheetIndex>0){
        for(int i=0;i<sheetIndex;i++){
            Sheet sheet = wb.getSheetAt(i);
            if(!EXCEL_HIDE_SHEET_NAME.equals(sheet.getSheetName())){
                DataValidation data_validation_list = null;
                //省份选项添加验证数据
                for(int a=2;a<3002;a++){
                    data_validation_list = getDataValidationByFormula(HIDE_SHEET_NAME_PROVINCE,a,2);
                    sheet.addValidationData(data_validation_list);
                    //城市选项添加验证数据
                    data_validation_list = getDataValidationByFormula("INDIRECT(B"+a+")",a,3);
                    sheet.addValidationData(data_validation_list);
                    //性别添加验证数据
                    data_validation_list = getDataValidationByFormula(HIDE_SHEET_NAME_SEX,a,1);
                    sheet.addValidationData(data_validation_list);
                }
            }
        }
    }
}
    /**
     * 使用已定义的数据源方式设置一个数据验证
     * @param formulaString
     * @param naturalRowIndex
     * @param naturalColumnIndex
     * @return
     */
    private static DataValidation getDataValidationByFormula(String formulaString,int naturalRowIndex,int naturalColumnIndex){
        //加载下拉列表内容
        DVConstraint constraint = DVConstraint.createFormulaListConstraint(formulaString);
        //设置数据有效性加载在哪个单元格上。
        //四个参数分别是：起始行、终止行、起始列、终止列
        int firstRow = naturalRowIndex-1;
        int lastRow = naturalRowIndex-1;
        int firstCol = naturalColumnIndex-1;
        int lastCol = naturalColumnIndex-1;
        CellRangeAddressList regions=new CellRangeAddressList(firstRow,lastRow,firstCol,lastCol);
        //数据有效性对象
        DataValidation data_validation_list = new HSSFDataValidation(regions,constraint);
        //设置输入信息提示信息
        data_validation_list.createPromptBox("下拉选择提示","请使用下拉方式选择合适的值！");
        //设置输入错误提示信息
        data_validation_list.createErrorBox("选择错误提示","你输入的值未在备选列表中，请下拉选择合适的值！");
        return data_validation_list;
    }
}