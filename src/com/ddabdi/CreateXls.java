package com.ddabdi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.Iterator;

/**
 * Created with IntelliJ IDEA.
 * User: deddy
 * Date: 12/3/17
 * Time: 1:08 PM
 * To change this template use File | Settings | File Templates.
 */
public class CreateXls {

   public static void genFile() throws IOException {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet1");

        HSSFRow row;
        HSSFCell cell1;
        HSSFCell cell2;

        // JUDUL
        HSSFCellStyle style= workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillBackgroundColor(HSSFFont.COLOR_RED);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        row = sheet.createRow(0);
        cell1 = row.createCell(0);
        cell1.setCellValue("Nama");
        cell1.setCellStyle(style);
        cell2 = row.createCell(1);
        cell2.setCellValue("Jumlah");
        cell2.setCellStyle(style);

        int i =1;
        while (i <=10){

            row = sheet.createRow(i);

            Barang barang = new Barang();
            barang.setNama("Barang [" + i + "]");
            barang.setJumlah(i);
            cell1 = row.createCell(0);
            cell1.setCellValue(barang.getNama());
            cell2 = row.createCell(1);
            cell2.setCellValue(barang.getJumlah());

            i++;
        }

        File file = new File("tesXls.xls");
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        fos.close();

        System.out.println("file created");

       }


    public static void readXls() throws IOException {
        File file = new File("tesXls.xls");
        FileInputStream fis = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(fis);
        HSSFSheet sheet = workbook.getSheet("sheet1");
        HSSFRow row ;
        HSSFCell cell;

        Iterator<Row> rows = sheet.iterator();

        while (rows.hasNext()){
            row = (HSSFRow) rows.next();
            Iterator<Cell> cells = row.cellIterator();

            while (cells.hasNext()){
                cell = (HSSFCell) cells.next();
                switch (cell.getCellTypeEnum()){
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                }
            }
            System.out.println();

        }


    }
}


class Barang {
    private String nama;
    private int jumlah;

    String getNama() {
        return nama;
    }

    void setNama(String nama) {
        this.nama = nama;
    }

    int getJumlah() {
        return jumlah;
    }

    void setJumlah(int jumlah) {
        this.jumlah = jumlah;
    }
}


