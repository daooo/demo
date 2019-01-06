package poi;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class ExcelReadUtil<T> {
    public static<T> List readExcel(String filePath, String classPath) {
        FileInputStream fis = null;
        Workbook workbook = null;
        List list = new ArrayList();
        try {
            fis = new FileInputStream(filePath);
            workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);
            //获取表头
            Row sheetHead = sheet.getRow(0);
            //通过反射获取实例对象的属性
            Class clazz = Class.forName(classPath);
            Field[] fields = clazz.getDeclaredFields();
            //外层循环遍历表的每一行
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                //创建一个实例对象用来封装表的一行数据
                T t = (T) clazz.newInstance();
                //获取表的一行
                Row row = sheet.getRow(i);
                //内层循环获取每一个单元格
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        //统一把数据设置为String类型
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                    }
                }
                //把一行数据封装进一个对象中
                setPropertity(row,fields,t,sheetHead);
                list.add(t);
            }
        }catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

    private static<T> void setPropertity(Row row, Field[] fields, T t, Row sheetHead) {
        //外层循环获得每一个对象的属性
        for (int i = 0; i < fields.length; i++) {
            //获取单元格数据
            String value = row.getCell(i).getStringCellValue();
            //获取属性名
            String fieldName = fields[i].getName();
            //获取属性类型，上面已经把单元格的数据都设置为了String，如果输入属性类型不是String，还要进行转换
            String type = fields[i].getGenericType().toString();
            //内层循环获取每一列表头的数据
            for (int j = 0; j < sheetHead.getLastCellNum(); j++) {
                //获取一列表头数据
                String rowName = sheetHead.getCell(j).getStringCellValue();
                //如果表头数据和属性名一致，就可以设置属性
                if (rowName.equals(fieldName)) {
                    try {
                        fields[i].setAccessible(true);
                        if (type.equals("class java.lang.Integer")) {
                            fields[i].set(t, Integer.parseInt(value));
                        }else{
                            fields[i].set(t, value);
                        }
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                    break;
                }
            }
        }
    }
}
