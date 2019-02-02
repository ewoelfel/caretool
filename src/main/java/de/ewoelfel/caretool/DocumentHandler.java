package de.ewoelfel.caretool;

import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.sql.Date;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;

import static java.util.stream.Stream.*;

@Component
public class DocumentHandler {

    private static final int SPECIALTY_COLUMN = 9;
    private static final String TPL_PREV_YEAR = "2017";
    private static final String TPL_CURRENT_YEAR = "2018";
    private static final Logger logger = LoggerFactory.getLogger(CmdRunner.class);

    private final InputStream templateAsStream = DocumentHandler.class.getResourceAsStream("/caretool_template.xlsx");

    private DateTimeFormatter monthDateFormatter = DateTimeFormatter.ofPattern("MMM yyyy", Locale.GERMANY);

    private GenerationContext context;
    private XSSFSheet tplMonthSheet;
    private XSSFSheet tplSummarySheet;
    private CreationHelper createHelper;
    private XSSFWorkbook exportWorkbook;


    public int generateDocuments(GenerationContext context) throws IOException {

        this.context = context;

        //read template workbook
        exportWorkbook = new XSSFWorkbook(templateAsStream);
        createHelper = exportWorkbook.getCreationHelper();
        this.tplSummarySheet = exportWorkbook.getSheetAt(0);
        this.tplMonthSheet = exportWorkbook.getSheetAt(1);

        createTabs();

        String firstFileName = String.format("%s Schichtplan %d.xlsx", context.getNames()[0], context.getYear().getValue());
            exportWorkbook(firstFileName);

            //copy the other ones
        of(context.getNames()).skip(1).forEach(name -> {
            String fileToCopyTo = String.format("%s Schichtplan %d.xlsx", name, context.getYear().getValue());
            try {
                Files.copy(Paths.get(firstFileName), Paths.get(fileToCopyTo), StandardCopyOption.REPLACE_EXISTING);
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        return context.getNames().length;
    }

    private void createTabs() {

        createSummaryTab();
        of(Month.values()).forEach(this::createMonthTab);
    }

    private void createMonthTab(Month month) {

        //set sheet name to e.g. Mrz 2019
        String sheetName = monthDateFormatter.format(LocalDate.of(context.getYear().getValue(), month, 1));
        XSSFSheet sheet = exportWorkbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = exportWorkbook.cloneSheet(1);
        }

        List<Day> daysOfMonth = context.getDaysInMonth().get(month);

        //copy all rows
        for (int row = 0; row < tplMonthSheet.getLastRowNum(); row++) {

            XSSFRow exportRow = sheet.getRow(row);
            XSSFRow tplRow = tplMonthSheet.getRow(row);

            boolean isDayInMonth = row > 0 && row <= daysOfMonth.size() && daysOfMonth.get(row - 1).getValue().getMonth() == month;
            Day day = row > 0 && row <= daysOfMonth.size() ? daysOfMonth.get(row - 1) : null;

            for (int col = 0; col < tplRow.getLastCellNum(); col++) {

                XSSFCell tplCell = tplRow.getCell(col);
                if (tplCell != null) {
                    XSSFCellStyle cellStyle = exportWorkbook.createCellStyle();
                    cellStyle.cloneStyleFrom(tplCell.getCellStyle());

                    if (row > 0 && row <= daysOfMonth.size()) {
                        setCellStyleByColumnNumForMonth(col, cellStyle);
                    }

                    XSSFCell cell = exportRow.getCell(col);

                    switch (tplCell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING:
                        case XSSFCell.CELL_TYPE_BLANK:
                            if (col == SPECIALTY_COLUMN && day != null) {
                                cell.setCellValue(day.isHolyday() ? "F" : "");
                            } else {
                                cell.setCellValue(tplCell.getStringCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if (col == 0 && isDayInMonth) {
                                cell.setCellValue(Date.valueOf(day.getValue()));
                            } else if (col == 0) {
                                cell.setCellValue("");
                            } else {
                                cell.setCellValue(tplCell.getNumericCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            if(row <= 31 && !isDayInMonth) {
                                cell.setCellType(Cell.CELL_TYPE_STRING);
                                cell.setCellValue("");
                            }
                            default:
                    }
                    cell.setCellStyle(cellStyle);
                }
            }
        }
        //set page setup to fit to one page width but multiple pages height
        sheet.getPrintSetup().setScale((short)60);
        sheet.getPrintSetup().setLandscape(true);
        sheet.getPrintSetup().setPaperSize(HSSFPrintSetup.A4_PAPERSIZE);
        exportWorkbook.getCTWorkbook().getSheets().getSheetArray(month.getValue()).setName(sheetName);
        logger.info("generated :" + month);
    }

    private void setCellStyleByColumnNumForMonth(int col, XSSFCellStyle cellStyle) {
        switch (col) {
            case 0:
                cellStyle.setDataFormat(
                        createHelper.createDataFormat().getFormat("dd.mm.yyyy"));
                break;
            case 1:
            case 2:
                cellStyle.setDataFormat(
                        createHelper.createDataFormat().getFormat("HH:mm"));
                break;
        }
    }

    private void createSummaryTab() {

        //copy all rows
        for (int row = 0; row < tplSummarySheet.getLastRowNum(); row++) {

            XSSFRow exportRow = tplSummarySheet.getRow(row);

            for (int col = 0; col < exportRow.getLastCellNum(); col++) {

                XSSFCell tplCell = exportRow.getCell(col);
                if (tplCell != null) {
                    XSSFCellStyle cellStyle = exportWorkbook.createCellStyle();
                    cellStyle.cloneStyleFrom(tplCell.getCellStyle());

                    //set Date header
                    boolean isDateRow = row == 7 && col > 0;
                    String sheetName = null;
                    LocalDate month = null;

                    if (col > 0 && col < 13) {
                        month = LocalDate.of(context.getYear().getValue(), Month.values()[col - 1], 1);
                        sheetName = monthDateFormatter.format(month);
                    }

                    XSSFCell cell = exportRow.getCell(col);

                    switch (tplCell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING:
                            String textToAdd;
                            textToAdd = tplCell.getStringCellValue()
                                    .replace(TPL_CURRENT_YEAR, context.getCurrentYear().getValue() + "")
                                    .replace(TPL_PREV_YEAR, context.getPreviousYear().getValue() + "");
                            cell.setCellValue(textToAdd);
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            if (row > 7 && row != 15 && row != 12 && col > 0) {
                                setSummaryFormulaByPosition(tplCell, cell, sheetName);

                            } else {
                                cell.setCellFormula(tplCell.getCellFormula());
                            }
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if (isDateRow) {
                                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MMM yyyy"));
                                cell.setCellValue(Date.valueOf(month));
                            } else {
                                cell.setCellValue(tplCell.getNumericCellValue());
                            }

                            break;
                        default:
                            cell.setCellValue(tplCell.getStringCellValue());
                    }
                    cell.setCellType(tplCell.getCellType());
                    cell.setCellStyle(cellStyle);
                }
            }
        }
        logger.info("generated :SUMMARY");

    }

    private void setSummaryFormulaByPosition(XSSFCell tplCell, XSSFCell cell, String sheetName) {
        int rowNum = tplCell.getRowIndex() >= 13 ? tplCell.getRowIndex() + 24 : tplCell.getRowIndex() + 25;
        cell.setCellFormula(String.format("'%s'!H%d", sheetName, rowNum));
    }

    private void exportWorkbook(String excelFileName) throws IOException {
        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        exportWorkbook.write(fileOut);
        fileOut.flush();
        fileOut.close();

    }
}
