/*
package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.xls.domain.StartPoint;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class XlsxCellsWritingServiceTest {

    private static final String SHEET_NAME = "One";

    @Test
    public void testWriteTopToBottom() {
        // Arrange
        List<String> values = Arrays.asList("val1", "val2", "val3", "val4", "val5");
        XlsxCellsWritingService xlsxCellsWritingService = new XlsxCellsWritingService();

        // Act
        xlsxCellsWritingService.writeTopToBottom(values, new StartPoint(3));

        // Assert
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(SHEET_NAME);
        Row firstRow = sheet.createRow(0);
        Cell firstRowFirstCell = firstRow.createCell(0);
        firstRowFirstCell.setCellValue("val1");
        Cell firstRowSecondCell = firstRow.createCell(1);
        firstRowSecondCell.setCellValue("val4");

        assertEquals(values.get(0), firstRowFirstCell.toString());
        assertEquals(values.get(3), firstRowSecondCell.toString());

        Row secondRow = sheet.createRow(1);
        Cell secondRowFirstCell = secondRow.createCell(0);
        secondRowFirstCell.setCellValue("val2");
        Cell secondRowSecondCell = secondRow.createCell(1);
        secondRowSecondCell.setCellValue("val5");

        assertEquals(values.get(1), secondRowFirstCell.toString());
        assertEquals(values.get(4), secondRowSecondCell.toString());

        Row thirdRow = sheet.createRow(2);
        Cell thirdRowFirstCell = thirdRow.createCell(0);
        thirdRowFirstCell.setCellValue("val3");

        assertEquals(values.get(2), thirdRowFirstCell.toString());
    }


    @Test
    public void testWriteTopToBottomJustOneRow() {
        // Arrange
        List<String> values = Arrays.asList("val1", "val2", "val3", "val4", "val5");
        XlsxCellsWritingService xlsxCellsWritingService = new XlsxCellsWritingService(SHEET_NAME);

        // Act
        xlsxCellsWritingService.writeTopToBottom(values, 1);

        // Assert
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(SHEET_NAME);
        Row firstRow = sheet.createRow(0);
        Cell firstRowFirstCell = firstRow.createCell(0);
        firstRowFirstCell.setCellValue("val1");
        Cell firstRowSecondCell = firstRow.createCell(1);
        firstRowSecondCell.setCellValue("val2");
        Cell firstRowLastCell = firstRow.createCell(4);
        firstRowLastCell.setCellValue("val5");

        assertEquals(values.get(0), firstRowFirstCell.toString());
        assertEquals(values.get(1), firstRowSecondCell.toString());
        assertEquals(values.get(4), firstRowLastCell.toString());
    }

    @Test
    public void testWriteTopToBottomJustOneRowAndMax3cols() {
        // Arrange
        List<String> values = Arrays.asList("val1", "val2", "val3", "val4", "val5");
        XlsxCellsWritingService xlsxCellsWritingService = new XlsxCellsWritingService(SHEET_NAME);

        // Act
        xlsxCellsWritingService.writeTopToBottom(values, 1, 3);

        // Assert
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(SHEET_NAME);
        Row firstRow = sheet.createRow(0);
        Cell firstRowFirstCell = firstRow.createCell(0);
        firstRowFirstCell.setCellValue("val1");
        Cell firstRowSecondCell = firstRow.createCell(1);
        firstRowSecondCell.setCellValue("val2");

        Row secondRow = sheet.createRow(1);
        Cell secondRowFirstCell = secondRow.createCell(0);
        secondRowFirstCell.setCellValue("val4");


        assertEquals(values.get(0), firstRowFirstCell.toString());
        assertEquals(values.get(1), firstRowSecondCell.toString());
        assertEquals(values.get(3), secondRowFirstCell.toString());
    }

    @Test
    public void testWriteTopToBottomWithMaxRowsAndMaxCols() {
        // Arrange
        List<String> values = Arrays.asList("val1", "val2", "val3", "val4", "val5", "val6", "val7", "val8", "val9");
        XlsxCellsWritingService xlsxCellsWritingService = new XlsxCellsWritingService(SHEET_NAME);

        // Act
        xlsxCellsWritingService.writeTopToBottom(values, 3, 3);

        // Assert
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(SHEET_NAME);
        Row firstRow = sheet.createRow(0);
        Cell firstRowFirstCell = firstRow.createCell(0);
        firstRowFirstCell.setCellValue("val1");
        Cell firstRowSecondCell = firstRow.createCell(1);
        firstRowSecondCell.setCellValue("val4");
        Cell firstRowThirdCell = firstRow.createCell(2);
        firstRowThirdCell.setCellValue("val7");

        assertEquals(values.get(0), firstRowFirstCell.toString());
        assertEquals(values.get(3), firstRowSecondCell.toString());
        assertEquals(values.get(6), firstRowThirdCell.toString());

        Row secondRow = sheet.createRow(1);
        Cell secondRowFirstCell = secondRow.createCell(0);
        secondRowFirstCell.setCellValue("val2");
        Cell secondRowSecondCell = secondRow.createCell(1);
        secondRowSecondCell.setCellValue("val5");
        Cell secondRowThirdCell = secondRow.createCell(2);
        secondRowThirdCell.setCellValue("val8");

        assertEquals(values.get(1), secondRowFirstCell.toString());
        assertEquals(values.get(4), secondRowSecondCell.toString());
        assertEquals(values.get(7), secondRowThirdCell.toString());

        Row thirdRow = sheet.createRow(2);
        Cell thirdRowFirstCell = thirdRow.createCell(0);
        thirdRowFirstCell.setCellValue("val3");
        Cell thirdRowSecondCell = thirdRow.createCell(1);
        thirdRowSecondCell.setCellValue("val6");
        Cell thirdRowThirdCell = thirdRow.createCell(2);
        thirdRowThirdCell.setCellValue("val9");

        assertEquals(values.get(2), thirdRowFirstCell.toString());
        assertEquals(values.get(5), thirdRowSecondCell.toString());
       assertEquals(values.get(8), thirdRowThirdCell.toString());
   }

*/
/*   @Test
   public void testWriteLeftToRight() {
       // Arrange
        List<String> values = Arrays.asList("valA", "valB", "valC", "valD", "valE");
        XlsxCellsWritingService xlsxCellsWritingService = new XlsxCellsWritingService(SHEET_NAME);

        // Act
        xlsxCellsWritingService.writeLeftToRight(values, 3);

        // Assert
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(SHEET_NAME);
        Row firstRow = sheet.createRow(0);
        Cell firstCell = firstRow.createCell(0);
        firstCell.setCellValue("valA");
        Cell secondCell = firstRow.createCell(1);
        secondCell.setCellValue("valB");
        Cell thirdCell = firstRow.createCell(2);
        thirdCell.setCellValue("valC");

        assertEquals(values.get(0), firstCell.toString());
        assertEquals(values.get(1), secondCell.toString());
        assertEquals(values.get(2), thirdCell.toString());

        Row secondRow = sheet.createRow(1);
        Cell fourthCell = secondRow.createCell(0);
        fourthCell.setCellValue("valD");
        Cell fifthCell = secondRow.createCell(1);
        fifthCell.setCellValue("valE");

        assertEquals(values.get(3), fourthCell.toString());
        assertEquals(values.get(4), fifthCell.toString());
   }*//*

   @Test
   public void testWriteTopToBottomDontExceedColIndex() {
       // Arrange
       List<String> values = Arrays.asList("valOne", "valTwo", "valThree", "valFour", "valFive", "valSix");
       XlsxCellsWritingService xlsxCellsWritingService = new XlsxCellsWritingService(SHEET_NAME);

       // Act
       xlsxCellsWritingService.writeTopToBottom(values, 2, 2);

       // Assert
       XSSFWorkbook workbook = new XSSFWorkbook();
       Sheet sheet = workbook.createSheet(SHEET_NAME);
       Row firstRow = sheet.createRow(0);
       Cell firstRowFirstCell = firstRow.createCell(0);
       firstRowFirstCell.setCellValue("valOne");
       Cell firstRowSecondCell = firstRow.createCell(1);
       firstRowSecondCell.setCellValue("valThree");

       assertEquals(values.get(0), firstRowFirstCell.toString());
       assertEquals(values.get(2), firstRowSecondCell.toString());

       Row secondRow = sheet.createRow(1);
       Cell secondRowFirstCell = secondRow.createCell(0);
       secondRowFirstCell.setCellValue("valTwo");
       Cell secondRowSecondCell = secondRow.createCell(1);
       secondRowSecondCell.setCellValue("valFour");

       assertEquals(values.get(1), secondRowFirstCell.toString());
       assertEquals(values.get(3), secondRowSecondCell.toString());

       Row thirdRow = sheet.createRow(2);
       Cell thirdRowFirstCell = thirdRow.createCell(0);
       thirdRowFirstCell.setCellValue("valFive");
       Cell thirdRowSecondCell = thirdRow.createCell(1);
       thirdRowSecondCell.setCellValue("valSix");

       assertEquals(values.get(4), thirdRowFirstCell.toString());
       assertEquals(values.get(5), thirdRowSecondCell.toString());

   }
}
*/
