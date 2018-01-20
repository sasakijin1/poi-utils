package demo;

import com.jin.commons.poi.OfficeIoResult;
import com.jin.commons.poi.OfficeIoUtils;
import com.jin.commons.poi.model.CellDataType;
import com.jin.commons.poi.model.CellOptions;
import com.jin.commons.poi.model.DatePattern;
import com.jin.commons.poi.model.SheetOptions;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class test {

    public static void main(String[] args) throws IOException, InvocationTargetException, IllegalAccessException {
        exportXlsTemplate();
        OfficeIoResult officeIoResult = importXls();
        List list = officeIoResult.getImportList();
        exportXls(list);
//        System.out.println(Integer.valueOf("0.0"));
    }

    private static OfficeIoResult importXls() throws InvocationTargetException, IllegalAccessException {

        SheetOptions sheet = new SheetOptions("测试1",TestDTO.class).addTitle("测试标题");

        sheet.setCellOptions(new CellOptions[]{
                new CellOptions("name","姓名"),
                new CellOptions("idCard","证件").addSubCells(
                        new CellOptions[]{
                                new CellOptions("idCardOne","证件1").addCellDataType(CellDataType.VARCHAR),
                                new CellOptions("idCardTwo","证件2")
                        }
                ),
                new CellOptions("phone","手机号").addCellDataType(CellDataType.VARCHAR),
                new CellOptions("birthday","出生年月").addCellDataType(CellDataType.DATE).addPattern(DatePattern.DATE_FORMAT_DAY),
                new CellOptions("date","年月").addSubCells(
                        new CellOptions[]{
                                new CellOptions("yearMonthStr","年月中文"),
                                new CellOptions("yearMonth","年月数字").addCellDataType(CellDataType.VARCHAR),
                                new CellOptions("dayStr","字符日期"),
                                new CellOptions("dayNic","数字日期"),
                                new CellOptions("dayFormat","格式日期")
                        }
                ),
                new CellOptions("status","状态").isSelect(),
                new CellOptions("select","下拉").isSelect().setSelectBind("key","status"),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER)
        });

        return OfficeIoUtils.importXlsx(new File("d:\\success.xlsx"),new SheetOptions[]{sheet});
    }

    private static void exportXlsTemplate() throws IOException {

        Map transactStatusFixed = new HashMap();
        transactStatusFixed.put("0", "未办理");
        transactStatusFixed.put("1", "办理中");
        transactStatusFixed.put("2", "已办理");
        transactStatusFixed.put("3", "已退回");

        List<Map<String,Object>> selectList = new ArrayList<Map<String, Object>>();
        Map<String,Object> aa = new HashMap();
        aa.put("id", "100");
        aa.put("name", "一灵灵");
        aa.put("key", "0");
        Map<String,Object> bb = new HashMap();
        bb.put("id", "101");
        bb.put("name", "一灵一");
        bb.put("key", "0");
        selectList.add(aa);
        selectList.add(bb);

        Map<String,Object> cc = new HashMap();
        cc.put("id", "200");
        cc.put("name", "2灵灵");
        cc.put("key", "1");
        Map<String,Object> dd = new HashMap();
        dd.put("id", "201");
        dd.put("name", "2灵2");
        dd.put("key", "1");
        selectList.add(cc);
        selectList.add(dd);

        Map<String,Object> ee = new HashMap();
        ee.put("id", "300");
        ee.put("name", "3灵灵");
        ee.put("key", "2");
        Map<String,Object> ff = new HashMap();
        ff.put("id", "301");
        ff.put("name", "3灵一");
        ff.put("key", "2");
        selectList.add(ee);
        selectList.add(ff);

        Map<String,Object> gg = new HashMap();
        gg.put("id", "400");
        gg.put("name", "4灵灵");
        gg.put("key", "3");
        Map<String,Object> hh = new HashMap();
        hh.put("id", "401");
        hh.put("name", "4灵一");
        hh.put("key", "3");
        selectList.add(hh);
        selectList.add(gg);

        SheetOptions sheet = new SheetOptions("测试1").addTitle("测试标题");
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
                new CellOptions("status","状态").addCellSelect(transactStatusFixed),
                new CellOptions("select","下拉").addCellSelect("id","name",selectList).setSelectBind("key","status"),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER)
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

        SheetOptions sheet = new SheetOptions("测试1",TestDTO.class).addTitle("测试标题");
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
                new CellOptions("status","状态"),
                new CellOptions("select","下拉"),
                new CellOptions("amountStr","金额文字"),
                new CellOptions("amountNum","金额数字").addCellDataType(CellDataType.NUMBER)
        });
        sheet.setExportData(list);

        FileOutputStream out = new FileOutputStream("d:\\successList.xlsx");
        OfficeIoResult result = OfficeIoUtils.exportXlsx(new SheetOptions[]{sheet});
        result.getResultWorkbook().write(out);
        out.close();
    }
}
