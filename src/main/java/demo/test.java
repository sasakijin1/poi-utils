package demo;

import poi.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class test {

    public static void main(String[] args) throws IOException, InvocationTargetException, IllegalAccessException {
        List list = importXls();
        exportXls(list);
        exportXlsTemplate(list);
    }

    private static List importXls() throws FileNotFoundException, InvocationTargetException, IllegalAccessException {

        SheetOptions sheet = new SheetOptions("测试1", 0,1,TestDTO.class);

        Map transactStatusFixed = new HashMap();
        transactStatusFixed.put("0", "未办理");
        transactStatusFixed.put("1", "办理中");
        transactStatusFixed.put("2", "已办理");
        transactStatusFixed.put("3", "已退回");

        sheet.setCellOptions(new CellOptions[]{
                new CellOptions("name","姓名"),
                new CellOptions("idCardOne","证件1").addCellDataType(CellDataType.VARCHAR),
                new CellOptions("idCardTwo","证件2"),
                new CellOptions("phone","手机号").addCellDataType(CellDataType.VARCHAR),
                new CellOptions("birthday","出生年月"),
                new CellOptions("yearMonthStr","年月中文"),
                new CellOptions("yearMonth","年月数字").addCellDataType(CellDataType.VARCHAR),
                new CellOptions("dayStr","字符日期"),
                new CellOptions("dayNic","数字日期"),
                new CellOptions("dayFormat","格式日期"),
                new CellOptions("select","下拉").addCellSelect(transactStatusFixed),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER),
                new CellOptions("status","状态"),
                new CellOptions("timestamp","timestamp").addCellDataType(CellDataType.TIMESTAMP)

        });

        OfficeIoResult officeIoResult = OfficeIoUtils.importXlsx(new File("d:\\success.xls"),new SheetOptions[]{sheet});
        return officeIoResult.getImportList();
    }

    private static void exportXlsTemplate(List list) throws IOException {

        Map transactStatusFixed = new HashMap();
        transactStatusFixed.put("0", "未办理");
        transactStatusFixed.put("1", "办理中");
        transactStatusFixed.put("2", "已办理");
        transactStatusFixed.put("3", "已退回");

        SheetOptions sheet = new SheetOptions("测试1", 0,1);
        sheet.setCellOptions(new CellOptions[]{
                new CellOptions("name","姓名"),
                new CellOptions("idCardOne","证件1"),
                new CellOptions("idCardTwo","证件2"),
                new CellOptions("phone","手机号"),
                new CellOptions("birthday","出生年月"),
                new CellOptions("yearMonthStr","年月中文"),
                new CellOptions("yearMonth","年月数字"),
                new CellOptions("dayStr","字符日期"),
                new CellOptions("dayNic","数字日期"),
                new CellOptions("dayFormat","格式日期"),
                new CellOptions("select","下拉").addCellSelect(transactStatusFixed),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER),
                new CellOptions("status","状态").addFixedMap(transactStatusFixed),
                new CellOptions("timestamp","timestamp").addCellDataType(CellDataType.TIMESTAMP)
        });

        OfficeIoResult result = OfficeIoUtils.exportXlsxTempalet(sheet);
        FileOutputStream out = new FileOutputStream("d:\\successTemplate.xlsx");
        result.getResultWorkbook().write(out);
        out.close();
    }

    private static void exportXls(List list) throws IOException {

        Map transactStatusFixed = new HashMap();
        transactStatusFixed.put("0", "未办理");
        transactStatusFixed.put("1", "办理中");
        transactStatusFixed.put("2", "已办理");
        transactStatusFixed.put("3", "已退回");

        SheetOptions sheet = new SheetOptions("测试1", 0,1);
        sheet.setCellOptions(new CellOptions[]{
                new CellOptions("name","姓名"),
                new CellOptions("idCardOne","证件1"),
                new CellOptions("idCardTwo","证件2"),
                new CellOptions("phone","手机号"),
                new CellOptions("birthday","出生年月").addCellDataType(CellDataType.DATE),
                new CellOptions("yearMonthStr","年月中文"),
                new CellOptions("yearMonth","年月数字"),
                new CellOptions("dayStr","字符日期"),
                new CellOptions("dayNic","数字日期"),
                new CellOptions("dayFormat","格式日期"),
                new CellOptions("select","下拉").addCellSelect(transactStatusFixed),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER),
                new CellOptions("status","状态").addFixedMap(transactStatusFixed),
                new CellOptions("timestamp","timestamp").addCellDataType(CellDataType.TIMESTAMP)
        });
        sheet.setExportData(list);

        FileOutputStream out = new FileOutputStream("d:\\successList.xlsx");
        OfficeIoResult result = OfficeIoUtils.exportXlsx(new SheetOptions[]{sheet});
        result.getResultWorkbook().write(out);
        out.close();
    }
}
