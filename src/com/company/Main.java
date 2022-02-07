package com.company;

import java.io.*;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {

    public static void main(String[] args) throws IOException {
        final File folder = new File("F:\\NN\\");

        final File exported = new File("F:\\test.txt");
        listFilesForFolder(folder, exported, 1);


    }

    public static void listFilesForFolder(final File folder, final File exported, int skipLine) throws IOException {
        FileOutputStream fileOut = null;
        XSSFWorkbook xssfWorkbook = null;
        try {
            int incr = 0;
            int col = 0;
            xssfWorkbook = new XSSFWorkbook();
            XSSFSheet createdSheet = xssfWorkbook.createSheet("text");

            for (final File fileEntry : folder.listFiles()) {

                String ext = fileEntry.getName().substring(fileEntry.getName().lastIndexOf('.') + 1);
                StringBuilder stringBuilder = new StringBuilder("");
                System.out.println("-->> " + fileEntry.getName() + "-->>");
                 fileOut = new FileOutputStream("F:\\test.xlsx");
                XSSFRow rowCreated = createdSheet.createRow((short)incr);

                InputStream file = new FileInputStream(fileEntry.getPath());
                Workbook workbook = StreamingReader.builder()
                        .rowCacheSize(10)    // number of rows to keep in memory (defaults to 10)
                        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                        .open(file);            // InputStream or File for XLSX file (required)
                Sheet sheetx = workbook.getSheetAt(0);
                //--->Create the sheet here


                Iterator<Row> rowIteratorx = sheetx.iterator();
                if (fileEntry.isDirectory()) {
                    continue;
                } else if (ext.equals("xlsx")) {
                   Row row = rowIteratorx.next();
                    row = rowIteratorx.next();
                    row = rowIteratorx.next();

                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {

                        Cell cell = cellIterator.next();
                        switch (cell.getCellType()) {
                            case STRING:
                                rowCreated.createCell(col).setCellValue(cell.getStringCellValue());
                                col++;
                                //System.out.print(cell.getStringCellValue() + "\t");
                                //stringBuilder.append(cell.getStringCellValue() + "\t");
                                break;
                            case NUMERIC:
                                //System.out.print(cell.getNumericCellValue() + "\t");
                                //stringBuilder.append(cell.getNumericCellValue() + "\t");
                                rowCreated.createCell(col).setCellValue(cell.getNumericCellValue());
                                col++;
                                break;
                            default:
                        }
                    }
                    xssfWorkbook.write(fileOut);

                    incr++;
                    col = 0;
                    //System.out.println( stringBuilder);
                    //System.out.println("");
                    workbook.close();
                    file.close();

                }
            }

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }finally {
            fileOut.close();
            xssfWorkbook.close();

        }
    }



}





