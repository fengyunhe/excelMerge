package org.example;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;

import java.io.File;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        ExcelWriterBuilder write = EasyExcel.write("1.xlsx");
        ExcelWriterSheetBuilder sheet = write.sheet(0);
        List<Object> allData = new ArrayList<>();
        for (File file : Objects.requireNonNull(new File("C:\\Users\\yy\\Downloads\\WeChat Files\\yangyan890328\\FileStorage\\File\\2022-11\\东方大厦1区（5-25层）人员信息登记\\东方大厦1区（5-25层）人员信息登记").listFiles())) {
            if (file.getName().contains("汇总表")) {
                continue;
            }

            allData.addAll(Arrays.asList(Collections.singletonMap(0,file.getName().replace(".xls", ""))));
            List<Object> sheetData = EasyExcelFactory.read(file).sheet(1).doReadSync();
            allData.addAll(sheetData);
            allData.addAll(Arrays.asList(Collections.singletonMap(0,"")));
        }
        sheet.doWrite(allData);
        sheet.build();
    }
}