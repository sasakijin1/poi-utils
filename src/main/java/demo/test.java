package demo;

import poi.*;
import poi.model.CellDataType;
import poi.model.CellOptions;
import poi.model.SheetOptions;
import poi.utils.FieldUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class test {

    public static void main(String[] args) throws IOException, InvocationTargetException, IllegalAccessException {
        List list = importXls();
        System.out.println(list.size());
        exportXls(list);
//        exportXlsTemplate();


    }

    private static List importXls() throws FileNotFoundException, InvocationTargetException, IllegalAccessException {

        SheetOptions sheet = new SheetOptions("测试1",TestDTO.class);

        Map transactStatusFixed = new HashMap();
        transactStatusFixed.put("0", "未办理");
        transactStatusFixed.put("1", "办理中");
        transactStatusFixed.put("2", "已办理");
        transactStatusFixed.put("3", "已退回");

        sheet.setCellOptions(new CellOptions[]{
                new CellOptions("name","姓名"),
                new CellOptions("idCard","证件").addSubCells(
                        new CellOptions[]{
                                new CellOptions("idCardOne","证件1").addCellDataType(CellDataType.VARCHAR),
                                new CellOptions("idCardTwo","证件2")
                        }
                ),
                new CellOptions("phone","手机号").addCellDataType(CellDataType.VARCHAR),
                new CellOptions("birthday","出生年月").addCellDataType(CellDataType.DATE),
                new CellOptions("date","年月").addSubCells(
                        new CellOptions[]{
                                new CellOptions("yearMonthStr","年月中文"),
                                new CellOptions("yearMonth","年月数字").addCellDataType(CellDataType.VARCHAR),
                                new CellOptions("dayStr","字符日期"),
                                new CellOptions("dayNic","数字日期"),
                                new CellOptions("dayFormat","格式日期")
                        }
                ),
                new CellOptions("select","下拉"),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER),
                new CellOptions("status","状态")
        });

        OfficeIoResult officeIoResult = OfficeIoUtils.importXlsx(new File("d:\\success.xlsx"),new SheetOptions[]{sheet});
        return officeIoResult.getImportList();
    }

    private static void exportXlsTemplate() throws IOException {

        Map transactStatusFixed = new HashMap();
        transactStatusFixed.put("0", "未办理");
        transactStatusFixed.put("1", "办理中");
        transactStatusFixed.put("2", "已办理");
        transactStatusFixed.put("3", "已退回");

        Map aa = new HashMap();
        aa.put("444", "4444");
        aa.put("555", "5555");
        aa.put("666", "6666");
        aa.put("777", "7777");

        SheetOptions sheet = new SheetOptions("测试1");
        sheet.setCellOptions(new CellOptions[]{
                new CellOptions("name","姓名"),
                new CellOptions("idCard","证件").addSubCells(
                        new CellOptions[]{
                            new CellOptions("idCardOne","证件1").addCellDataType(CellDataType.VARCHAR),
                            new CellOptions("idCardTwo","证件2")
                        }
                ),
                new CellOptions("phone","手机号").addCellDataType(CellDataType.VARCHAR),
                new CellOptions("birthday","出生年月").addCellDataType(CellDataType.DATE),
                new CellOptions("date","年月").addSubCells(
                        new CellOptions[]{
                            new CellOptions("yearMonthStr","年月中文"),
                            new CellOptions("yearMonth","年月数字").addCellDataType(CellDataType.VARCHAR),
                            new CellOptions("dayStr","字符日期"),
                            new CellOptions("dayNic","数字日期"),
                            new CellOptions("dayFormat","格式日期")
                        }
                ),
                new CellOptions("select","下拉").addCellSelect(null),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER),
                new CellOptions("status","状态").addCellSelect(transactStatusFixed)
        });

        OfficeIoResult result = OfficeIoUtils.exportXlsxTemplate(sheet);
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

        SheetOptions sheet = new SheetOptions("测试1");
        sheet.setCellOptions(new CellOptions[]{
                new CellOptions("name","姓名"),
                new CellOptions("idCard","证件").addSubCells(
                        new CellOptions[]{
                                new CellOptions("idCardOne","证件1").addCellDataType(CellDataType.VARCHAR),
                                new CellOptions("idCardTwo","证件2")
                        }
                ),
                new CellOptions("phone","手机号").addCellDataType(CellDataType.VARCHAR),
                new CellOptions("birthday","出生年月").addCellDataType(CellDataType.DATE),
                new CellOptions("date","年月").addSubCells(
                        new CellOptions[]{
                                new CellOptions("yearMonthStr","年月中文"),
                                new CellOptions("yearMonth","年月数字").addCellDataType(CellDataType.VARCHAR),
                                new CellOptions("dayStr","字符日期"),
                                new CellOptions("dayNic","数字日期"),
                                new CellOptions("dayFormat","格式日期")
                        }
                ),
                new CellOptions("select","下拉"),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER),
                new CellOptions("status","状态")
        });
        sheet.setExportData(list);

        FileOutputStream out = new FileOutputStream("d:\\successList.xlsx");
        OfficeIoResult result = OfficeIoUtils.exportXlsx(new SheetOptions[]{sheet});
        result.getResultWorkbook().write(out);
        out.close();
    }
}
