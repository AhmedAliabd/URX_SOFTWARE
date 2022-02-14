package com.company;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Iterator;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {
    //sorted = Arrays.stream(sorted).filter(s -> !s.getName().split("-")[1].equalsIgnoreCase(" Trash.xlsx")).toArray(File[]::new);

    public static void main(String[] args) throws IOException {
        final File folder = new File("F:\\NN\\");
        File[] sorted = sortFiles(folder);

        final File exported = new File("F:\\test.txt");
        listFilesForFolder(sorted, exported, 1);
    }

    public static void listFilesForFolder(final File[] folder, final File exported, int skipLine) throws IOException {
        FileOutputStream fileOut = null;
        XSSFWorkbook xssfWorkbook = null;
        try {
            int incr = 0;
            int col = 0;

            //--->Create the sheet to export the data
            xssfWorkbook = new XSSFWorkbook();
            XSSFSheet createdSheet = xssfWorkbook.createSheet("text");
            //---> loop throw the files in the selected folder
            for (final File fileEntry : folder) {
                //Get the file extension
                String ext = fileEntry.getName().substring(fileEntry.getName().lastIndexOf('.') + 1);
                System.out.println("-->> " + fileEntry.getName() + "-->>");
                //---> Create the file to export the data
                fileOut = new FileOutputStream("F:\\test.xlsx");
                //---> Create the row to insert the cells in
                XSSFRow rowCreated = createdSheet.createRow((short)incr);
                //Read the file selected by the loop
                InputStream file = new FileInputStream(fileEntry.getPath());
                // The following block of code is by the (Excel Streaming Reader)
                // https://github.com/monitorjbl/excel-streaming-reader
                Workbook workbook = StreamingReader.builder()
                        .rowCacheSize(10)    // number of rows to keep in memory (defaults to 10)
                        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                        .open(file);            // InputStream or File for XLSX file (required)

                Sheet sheetx = workbook.getSheetAt(0);
                Iterator<Row> rowIteratorx = sheetx.iterator();

                if (fileEntry.isDirectory()) {
                    continue;
                } else if (ext.equals("xlsx")) {
                    //Skip the head row
                    Row row = rowIteratorx.next();
                    //Skip the specified rows
                    for(int j = 0; j < skipLine; j++)
                    {
                        row = rowIteratorx.next();

                    }
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        switch (cell.getCellType()) {
                            case STRING:
                                rowCreated.createCell(col).setCellValue(cell.getStringCellValue());
                                col++;
                                break;
                            case NUMERIC:
                                rowCreated.createCell(col).setCellValue(cell.getNumericCellValue());
                                col++;
                                break;
                            case BLANK:
                                rowCreated.createCell(col).setCellValue(cell.getStringCellValue());
                                col++;
                                break;
                            case BOOLEAN:
                                rowCreated.createCell(col).setCellValue(cell.getBooleanCellValue());
                                col++;
                                break;
                            case ERROR:
                                rowCreated.createCell(col).setCellValue(cell.getErrorCellValue());
                                col++;
                                break;
                            case FORMULA:
                                rowCreated.createCell(col).setCellValue(cell.getCellFormula());
                                col++;
                                break;
                            default:
                        }
                    }
                    xssfWorkbook.write(fileOut);
                    incr++;
                    col = 0;
                    workbook.close();//Close the current worksheet
                    file.close();//Close the current file
                }
            }

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }finally {
            fileOut.close();
            xssfWorkbook.close();
        }
    }

    public static File[] sortFiles(File file)
    {
        File[] xx = file.listFiles();
        Arrays.sort(xx, new Comparator<File>() {
            @Override
            public int compare(File o1, File o2) {
                int n1 = extractNumber(o1.getName());
                int n2 = extractNumber(o2.getName());
                return n1 - n2;
            }
            private int extractNumber(String name) {
                int i = 0;
                try {
                    //NEED to check for index problem here !!!!!!!
                    //"(?<=\\D)(?=\\d)"
                    String part = name.split("(\\D+)")[1]; //This step can be reduced by using regex only without split method

                    //String xx = part.split("[.]")[0];
                    i = Integer.parseInt(part);
                } catch(Exception e) {
                    i = 0; // if filename does not match the format
                    // then default to 0
                }

                return i;
            }
        });
        return xx;
    }

}





