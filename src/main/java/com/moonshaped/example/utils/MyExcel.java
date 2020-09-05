package com.moonshaped.example.utils;

import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class MyExcel {

	private String pathname;
    private Workbook workbook;
    private Sheet sheet1;

    /**使用栗子
     * WriteExcel excel = new WriteExcel("D:\\myexcel.xlsx");
     * excel.write(new String[]{"1","2"}, 0);//在第1行第1個單元格寫入1,第一行第二個單元格寫入2
     */
    public void write(String[] writeStrings, int rowNumber) throws Exception {
        //將內容寫入指定的行號中
        Row row = sheet1.createRow(rowNumber);
        //遍歷整行中的列序號
        for (int j = 0; j < writeStrings.length; j++) {
            //根據行指定列座標j,然後在單元格中寫入資料
            Cell cell = row.createCell(j);
            cell.setCellValue(writeStrings[j]);
        }
        OutputStream stream = new FileOutputStream(pathname);
        workbook.write(stream);
        stream.close();
    }

    public MyExcel(String excelPath) throws Exception {
        //在excelPath中需要指定具體的檔名(需要帶上.xls或.xlsx的字尾)
        this.pathname = excelPath;
        String fileType = excelPath.substring(excelPath.lastIndexOf(".") + 1, excelPath.length());
        //建立文件物件
        if (fileType.equals("xls")) {
            //如果是.xls,就new HSSFWorkbook()
            workbook = new HSSFWorkbook();
        } else {
            throw new Exception("文件格式字尾不正確!!！");
        }
        // 建立表sheet
        sheet1 = workbook.createSheet("sheet1");
    }
}
