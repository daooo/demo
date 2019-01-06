package poi;

import java.util.List;

public class Test {
    public static void main(String[] args) {
        List list = ExcelReadUtil.<User>readExcel("h:/user.xlsx","poi.User");
        System.out.print(list);
    }
}
