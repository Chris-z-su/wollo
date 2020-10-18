package com.javase.export;


import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;

public class RowHeightTest {
    public static void main(String[] args) throws Exception {
        //设置行高：https://blog.csdn.net/lipinganq/article/details/78081300
        //单元格数据的对齐方式：https://blog.csdn.net/LinBilin_/article/details/54375262
        String pathName = "C:\\Users\\Administrator\\Desktop\\test.xls";

        File file = new File(pathName);
        if (file.exists()) {
            file.delete();
        }
        BufferedOutputStream out = null;
        try {
            out = new BufferedOutputStream(new FileOutputStream(pathName));
            exportExcel(out);
        } finally {
            out.close();
        }
    }

    private static void exportExcel(BufferedOutputStream out) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 格式化Sheet名，使其合法
        String safeSheetName = WorkbookUtil.createSafeSheetName("日历");
        HSSFSheet sheet = workbook.createSheet(safeSheetName);
        sheet.setColumnWidth(0, 100*256);
        HSSFRow row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("行高16px = 12pt = 240twips");
        // 单位为twips(缇)
        row3.setHeight((short)(12*20));

        HSSFRow row5 = sheet.createRow(5);
        row5.createCell(0).setCellValue("行高32px = 24pt = 480twips");
        row5.setHeight((short)(24*20));

        HSSFRow row7 = sheet.createRow(7);
        row7.createCell(0).setCellValue("行高48px = 36pt = 720twips");
        // 单位为pt(磅)
        row7.setHeightInPoints(36);

        HSSFRow row9 = sheet.createRow(9);
        row9.createCell(0).setCellValue("行高64px = 48pt = 1440twips");
        // 单位为pt(磅)
        row9.setHeightInPoints(48);
        workbook.write(out);
    }
}
