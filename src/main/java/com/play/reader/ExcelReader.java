package com.play.reader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

@RestController
public class ExcelReader {

    @RequestMapping(value = "/test1", method = RequestMethod.GET)
    public void processReader() throws IOException, FileNotFoundException {

        final String FILE_NAME = "D:\\IntelliJ_PlayStation\\POI_Demo\\target\\ExceptionAudit1.xlsx";

        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Sheet datatypeSheet1 = workbook.getSheetAt(0);
        Row rowIterator = datatypeSheet.getRow(0);
        Iterator<Row> subIterator = datatypeSheet1.iterator();
        int incrementer = 0;
        int horizontalInc = 0;
        while (rowIterator != null) {
  //          Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = rowIterator.iterator();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                if (currentCell.getColumnIndex() > 3) {
                    incrementer ++;
                    //getCellTypeEnum shown as deprecated for version 3.15
                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
  //                      System.out.print(currentCell.getStringCellValue() + "  ");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
//                        System.out.print(currentCell.getNumericCellValue() + "  ");
                    }
                    String upc = currentCell.getStringCellValue();
 //                   System.out.print(upc + ",");
 //                   String contentAppend = upc;
                    while (subIterator.hasNext()) {
                        String contentAppend = upc;
                        Row currentRow = subIterator.next();
                        Iterator<Cell> subCellIterator = currentRow.iterator();
                        while (subCellIterator.hasNext()) {

                            Cell subCurrentCell = subCellIterator.next();
                            String value = null;
                            if (currentRow.getRowNum() > 0 && subCurrentCell.getColumnIndex() < 4) {

                                //getCellTypeEnum shown as deprecated for version 3.15
                                //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                                if (subCurrentCell.getCellTypeEnum() == CellType.STRING) {
  //                                  System.out.print(subCurrentCell.getStringCellValue() + "--");
                                    value = subCurrentCell.getStringCellValue();
                                } else if (subCurrentCell.getCellTypeEnum() == CellType.NUMERIC) {
  //                                  System.out.print(subCurrentCell.getNumericCellValue() + "--");
                                       double value1 = subCurrentCell.getNumericCellValue();
                                     value =Double. toString(value1);
                                }
                            }


                            if (currentRow.getRowNum() > 0 && subCurrentCell.getColumnIndex() == 3) {
                                contentAppend = contentAppend.concat(",").concat(value);
                                horizontalInc = subCurrentCell.getColumnIndex() + incrementer;
                                if (currentRow.getCell(horizontalInc).getCellTypeEnum() == CellType.STRING) {
                                    //                                  System.out.print(subCurrentCell.getStringCellValue() + "--");
                                    value = currentRow.getCell(horizontalInc).getStringCellValue();
                                } else if (currentRow.getCell(horizontalInc).getCellTypeEnum() == CellType.NUMERIC) {
                                    //                                  System.out.print(subCurrentCell.getNumericCellValue() + "--");
                                    double value1 = currentRow.getCell(horizontalInc).getNumericCellValue();
                                    value = Double. toString(value1);
                                }
                            }

                            if (currentRow.getRowNum() > 0 && subCurrentCell.getColumnIndex() < 4) {
                                contentAppend = contentAppend.concat(",").concat(value);
                            }
                        }
                        if (currentRow.getRowNum() > 0 ) {
                            System.out.print(contentAppend);
                            System.out.print("\n");
                        }
                        contentAppend = " ";
                    }
                    subIterator = datatypeSheet1.iterator();
                }
            }

            System.out.println();
            rowIterator = null;
        }
    }
}

