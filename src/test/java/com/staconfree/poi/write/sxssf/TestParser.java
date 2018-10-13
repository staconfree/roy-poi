package com.staconfree.poi.write.sxssf;

import com.staconfree.clazz.model.Student;
import com.staconfree.clazz.model.Teacher;
import com.staconfree.poi.write.WriteCallBack;
import com.staconfree.poi.write.common.SimpleWriteCellHelper;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2018/10/13 0013.
 */
public class TestParser {


    public static void main(String[] args) {
        long bgn = System.currentTimeMillis();
        Map<String, Object> dataStack = new HashMap<String, Object>();
        dataStack.put("orderTime1", "2018-10-01");
        dataStack.put("orderTime2", "2018-10-03");
        List<Model> userList = initUserList();
        dataStack.put("userList", userList);
        List<Model> brandList = initBrandList();
        dataStack.put("brandList", brandList);
        String classpath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
        String fromFilePath = classpath + "/template.xlsx";
        System.out.println(fromFilePath);
        SxssfWriteCellHelper cellHelper = new SxssfWriteCellHelper();
        try {
            SxssfExcelParser excelParse = null;
            excelParse = new SxssfExcelParser(fromFilePath, classpath+"/test_parser.xlsx");
            WriteCallBack writeCallBack = new WriteCallBack() {
                @Override
                public void changeRow(Row row, Object rowData, Sheet sheet) {
                }
                @Override
                public void changeCell(Cell cell, Object rowData, Sheet sheet) {
                    if (cell.getRowIndex() == 2) {
                        CellStyle cellStyle = cell.getCellStyle();
                        cellStyle.setFillForegroundColor(HSSFColor.YELLOW.index);
                        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    }
                    String cellValue = cellHelper.getCellValue(cell);
                    if (cell.getColumnIndex()==0&&cellValue!=null&&
                            cellValue.contains("合计")) {
                        CellStyle oldCellStyle = cell.getCellStyle();
                        try {
                            CellStyle newCellStyle = cellHelper.cloneCellStyle(oldCellStyle,sheet.getWorkbook());
                            cell.setCellStyle(newCellStyle);
                            newCellStyle.setFillForegroundColor(HSSFColor.YELLOW.index);
                            newCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(cell.getRowIndex(),cell.getRowIndex(),cell.getColumnIndex(),cell.getColumnIndex()+4);
                        sheet.addMergedRegion(cellRangeAddress);
                    }
                }
            };
            excelParse.setWriteCallBack(writeCallBack);
            excelParse.setTemplateSheetName("会员报表");
            excelParse.parse(dataStack);
            excelParse.endSheet();
            excelParse.setTemplateSheetName("机型报表");
            excelParse.parse(dataStack);
            excelParse.endSheet();
            excelParse.flush();
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("end:" + (System.currentTimeMillis() - bgn));
    }

    private static List<Model> initBrandList() {
        List<Model> list = new ArrayList<Model>();
        for (int i = 0; i < 5; i++) {
            if (i==0){
                Model model = new Model("1","华为","荣耀X1","代工厂1","20181001","1","1000.0");
                list.add(model);
            } else if (i==1){
                Model model = new Model("2","华为","荣耀X1","代工厂1","20181002","3","3000.0");
                list.add(model);
                model = new Model("华为合计","","","","","4","4000.0");
                list.add(model);
            } else if (i==2){
                Model model = new Model("3","小米","红米6","代工厂1","20181001","1","800.0");
                list.add(model);
            } else if (i==3){
                Model model = new Model("4","小米","红米6","代工厂2","20181001","2","1600.0");
                list.add(model);
                model = new Model("小米合计","","","","","3","2400.0");
                list.add(model);
            } else if (i==4){
                Model model = new Model("5","一加","一加6","代工厂1","20181001","1","3000.0");
                list.add(model);
                model = new Model("一加合计","","","","","1","3000.0");
                list.add(model);
            }
        }
        return list;
    }


    public static List<Model> initUserList() {
        List<Model> list = new ArrayList<Model>();
        for (int i = 0; i < 5; i++) {
            if (i==0){
                User user = new User("会员1","13800010001");
                Model model = new Model(user,"20181001","1","1000.0");
                list.add(model);
            } else if (i==1){
                User user = new User("会员1","13800010001");
                Model model = new Model(user,"20181002","3","3000.0");
                list.add(model);
            } else if (i==2){
                User user = new User("会员2","13800020002");
                Model model = new Model(user,"20181001","1","800.0");
                list.add(model);
            } else if (i==3){
                User user = new User("会员2","13800020002");
                Model model = new Model(user,"20181001","2","1600.0");
                list.add(model);
            } else if (i==4){
                User user = new User("会员3","13800030003");
                Model model = new Model(user,"20181001","1","3000.0");
                list.add(model);
            }
        }
        return list;
    }

    static class Model {
        private String id;
        private User user;
        private String brand;
        private String model;
        private String vmi;
        private String buyDate;
        private String num;
        private String price;

        public Model(){

        }

        public Model(String id,String brand,String model,String vmi,String buyDate,String num,String price){
            this.id=id;
            this.brand=brand;
            this.model=model;
            this.vmi=vmi;
            this.buyDate=buyDate;
            this.num=num;
            this.price=price;
        }


        public Model(User user,String buyDate,String num,String price){
            this.user=user;
            this.buyDate=buyDate;
            this.num=num;
            this.price=price;
        }

        public String getId() {
            return id;
        }

        public void setId(String id) {
            this.id = id;
        }

        public User getUser() {
            return user;
        }

        public void setUser(User user) {
            this.user = user;
        }

        public String getBrand() {
            return brand;
        }

        public void setBrand(String brand) {
            this.brand = brand;
        }

        public String getModel() {
            return model;
        }

        public void setModel(String model) {
            this.model = model;
        }

        public String getVmi() {
            return vmi;
        }

        public void setVmi(String vmi) {
            this.vmi = vmi;
        }

        public String getBuyDate() {
            return buyDate;
        }

        public void setBuyDate(String buyDate) {
            this.buyDate = buyDate;
        }

        public String getNum() {
            return num;
        }

        public void setNum(String num) {
            this.num = num;
        }

        public String getPrice() {
            return price;
        }

        public void setPrice(String price) {
            this.price = price;
        }
    }

    static class User {
        private String userCode;
        private String mobile;

        public User(){

        }
        public User(String userCode,String mobile){
            this.userCode=userCode;
            this.mobile=mobile;
        }

        public String getUserCode() {
            return userCode;
        }

        public void setUserCode(String userCode) {
            this.userCode = userCode;
        }

        public String getMobile() {
            return mobile;
        }

        public void setMobile(String mobile) {
            this.mobile = mobile;
        }
    }
}
