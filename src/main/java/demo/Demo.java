package demo;

import com.jin.commons.poi.OfficeIoResult;
import com.jin.commons.poi.OfficeIoUtils;
import com.jin.commons.poi.model.CellSettings;
import com.jin.commons.poi.model.DatePattern;
import com.jin.commons.poi.model.FormulaType;
import com.jin.commons.poi.model.SheetSettings;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Demo {

    public static void main(String[] args) throws IOException, NoSuchAlgorithmException {
        exportXlsTemplate();
        OfficeIoResult officeIoResult = importXls();

        List list = officeIoResult.getImportList();
        if (list != null && list.size() > 0){
            exportXls(list);
        }
        if (!officeIoResult.isCompleted()){
            OfficeIoResult result = OfficeIoUtils.exportErrorRecord(officeIoResult.getSheetSettings(),officeIoResult.getErrRecordRows());
            FileOutputStream out = new FileOutputStream("d:\\successError.xlsx");
            result.getResultWorkbook().write(out);
            out.close();
        }

    }

    private static OfficeIoResult importXls() {

        SheetSettings sheet = new SheetSettings("测试1",SocialDTO.class).addTitle("TEST");

        sheet.setCellSettings(new CellSettings[]{
            new CellSettings("empName","empName"),
            new CellSettings("companyId","companyId").isSelect(),
            new CellSettings("socialRuleId","socialRuleId").isSelect().setSelectBind("key","companyId"),
            new CellSettings("fundRuleId","fundRuleId").isSelect().setSelectBind("key","companyId"),
            new CellSettings("idCardType","idCardType").addStaticValue("1"),
            new CellSettings("idCard","idCard"),
            new CellSettings("inDate","inDate").addPattern(DatePattern.DATE_FORMAT_DAY),
            new CellSettings("phone","phone"),
            new CellSettings("status","status").isSelect(),
            new CellSettings("company","company").addSubCells(new CellSettings[]{
                new CellSettings("companyBase","companyBase"),
                new CellSettings("companyRatio","companyRatio"),
                new CellSettings("companyAmount","companyAmount")
            }),
            new CellSettings("personal","personal").addSubCells(new CellSettings[]{
                new CellSettings("personalBase","personalBase"),
                new CellSettings("personalRatio","personalRatio"),
                new CellSettings("personalAmount","personalAmount")
            })
        });

        OfficeIoResult officeIoResult = OfficeIoUtils.importXlsx(new File("d:\\success.xlsx"),new SheetSettings[]{sheet});
        return officeIoResult;
    }

    private static void exportXlsTemplate() throws IOException {

        Map transactStatusFixed = new HashMap();
        transactStatusFixed.put("0", "未办理");
        transactStatusFixed.put("1", "办理中");
        transactStatusFixed.put("2", "已办理");
        transactStatusFixed.put("3", "已退回");

        List companyList = new ArrayList();
        companyList.add("A0001");
        companyList.add("A0002");
        companyList.add("A0()&$%003");
        companyList.add("A0（你好）004");
        companyList.add("A0005");
        String[] companyArray = new String[companyList.size()];
        companyList.toArray(companyArray);

        List<Map<String,Object>> socialRule = new ArrayList<Map<String, Object>>();
        Map<String,Object> aa = new HashMap();
        aa.put("id", "100");
        aa.put("name", "一灵灵");
        aa.put("key", "A0001");
        Map<String,Object> bb = new HashMap();
        bb.put("id", "101");
        bb.put("name", "一灵一");
        bb.put("key", "A0001");
        socialRule.add(aa);
        socialRule.add(bb);

        Map<String,Object> cc = new HashMap();
        cc.put("id", "200");
        cc.put("name", "2灵灵");
        cc.put("key", "A0002");
        Map<String,Object> dd = new HashMap();
        dd.put("id", "201");
        dd.put("name", "2灵2");
        dd.put("key", "A0002");
        socialRule.add(cc);
        socialRule.add(dd);

        List<Map<String,Object>> fundRule = new ArrayList<Map<String, Object>>();
        Map<String,Object> ee = new HashMap();
        ee.put("id", "300");
        ee.put("name", "3灵灵");
        ee.put("key", "A0()&$%003");
        Map<String,Object> ff = new HashMap();
        ff.put("id", "301");
        ff.put("name", "3灵一");
        ff.put("key", "A0()&$%003");
        fundRule.add(ee);
        fundRule.add(ff);

        Map<String,Object> gg = new HashMap();
        gg.put("id", "400");
        gg.put("name", "4灵灵");
        gg.put("key", "A0（你好）004");
        Map<String,Object> hh = new HashMap();
        hh.put("id", "401");
        hh.put("name", "4灵一");
        hh.put("key", "A0005");
        fundRule.add(hh);
        fundRule.add(gg);

        SheetSettings sheet = new SheetSettings("测试1",SocialDTO.class).addTitle("TEST");
        sheet.setCellSettings(new CellSettings[]{
            new CellSettings("empName","empName"),
            new CellSettings("companyId","companyId").addCellSelect(companyArray),
            new CellSettings("socialRuleId","socialRuleId").addCellSelect("id","name",socialRule).setSelectBind("key","companyId"),
            new CellSettings("fundRuleId","fundRuleId").addCellSelect("id","name",fundRule).setSelectBind("key","companyId"),
            new CellSettings("idCardType","idCardType").addStaticValue("1"),
            new CellSettings("idCard","idCard"),
            new CellSettings("inDate","inDate").addPattern(DatePattern.DATE_FORMAT_DAY),
            new CellSettings("phone","phone"),
            new CellSettings("status","status").addCellSelect(transactStatusFixed),
            new CellSettings("company","company").addSubCells(new CellSettings[]{
                    new CellSettings("companyBase","companyBase"),
                    new CellSettings("companyRatio","companyRatio"),
                    new CellSettings("companyAmount","companyAmount").addFormulaGroupName("total")
            }),
            new CellSettings("personal","personal").addSubCells(new CellSettings[]{
                    new CellSettings("personalBase","personalBase"),
                    new CellSettings("personalRatio","personalRatio"),
                    new CellSettings("personalAmount","personalAmount").addFormulaGroupName("total")
            })
        });

        OfficeIoResult result = OfficeIoUtils.exportXlsxTemplate(sheet);
        FileOutputStream out = new FileOutputStream("d:\\successTemplate.xlsx");
        result.getResultWorkbook().write(out);
        out.close();
    }

    private static void exportXls(List list) throws IOException {

        SheetSettings sheet = new SheetSettings("测试1",SocialDTO.class).addTitle("TEST");
        sheet.setCellSettings(new CellSettings[]{
                new CellSettings("empName","empName"),
                new CellSettings("companyId","companyId"),
                new CellSettings("socialRuleId","socialRuleId"),
                new CellSettings("fundRuleId","fundRuleId"),
                new CellSettings("idCardType","idCardType"),
                new CellSettings("idCard","idCard"),
                new CellSettings("inDate","inDate").addPattern(DatePattern.DATE_FORMAT_DAY),
                new CellSettings("phone","phone"),
                new CellSettings("status","status").isSelect(),
                new CellSettings("company","company").addSubCells(new CellSettings[]{
                        new CellSettings("companyBase","companyBase"),
                        new CellSettings("companyRatio","companyRatio"),
                        new CellSettings("companyAmount","companyAmount").addFormulaGroupName("total")
                }),
                new CellSettings("personal","personal").addSubCells(new CellSettings[]{
                        new CellSettings("personalBase","personalBase"),
                        new CellSettings("personalRatio","personalRatio"),
                        new CellSettings("personalAmount","personalAmount").addFormulaGroupName("total")
                })
//                new CellSettings("total","total").addFormulaSettings(FormulaType.SUM,"total")
        });
        sheet.setExportData(list);

        FileOutputStream out = new FileOutputStream("d:\\successList.xlsx");
        OfficeIoResult result = OfficeIoUtils.exportXlsx(new SheetSettings[]{sheet});
        result.getResultWorkbook().write(out);
        out.close();
    }
}
