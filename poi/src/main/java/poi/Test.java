package poi;

import java.util.List;

public class Test {
    public static void main(String[] args) {
        //edit by lucy
        //edit by marry1
        //lucy branch
        //marry branch
        //lucy branch
        List list = ExcelReadUtil.<User>readExcel("h:/user.xlsx","poi.User");
        System.out.print(list);
    }
}
