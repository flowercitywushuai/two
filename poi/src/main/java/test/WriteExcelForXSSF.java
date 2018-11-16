package test;

import org.apache.poi.hssf.util.HSSFColor;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * 在这里对Excel写入内容使用的XSSF（HSSF也是同样的道理，不再多举例）：
 */

public class WriteExcelForXSSF {

    public void write() {
        //新建工作文档
        Workbook workbook = new XSSFWorkbook();
        //设置脚本
        Sheet sheet = workbook.createSheet("0");
        //设置行
        Row row = sheet.createRow(0);
        //获取样式对象
        CellStyle cellStyle = workbook.createCellStyle();
        // 设置这些样式
        cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        row.createCell(0).setCellStyle(cellStyle);
        row.createCell(0).setCellValue("姓名");
        row.createCell(1).setCellStyle(cellStyle);
        row.createCell(1).setCellValue("年龄");
        row.createCell(2).setCellStyle(cellStyle);
        row.createCell(2).setCellValue("电话");
        row.createCell(3).setCellStyle(cellStyle);
        row.createCell(3).setCellValue("住址");

        for (int i = 1; i <= 10; i++) {
            Row nrow = sheet.createRow(i);
            Cell ncell = nrow.createCell(0);
            ncell.setCellValue("1");
            ncell = nrow.createCell(1);
            ncell.setCellValue("2");
            ncell = nrow.createCell(2);
            ncell.setCellValue("3");
            ncell = nrow.createCell(3);
            ncell.setCellValue("4");
        }

        workbook.setSheetName(0, "信息");
        try {
            File file = new File("F:\\idea_workspace\\ideaworkspace11\\poi\\src\\main\\resources\\poi.xlsx");
            FileOutputStream fileoutputStream = new FileOutputStream(file);
            workbook.write(fileoutputStream);
            fileoutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        new WriteExcelForXSSF().write();
    }

}
