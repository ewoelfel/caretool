package de.ewoelfel.caretool;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.*;
import java.sql.Date;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.stream.Stream;

import static java.util.stream.Stream.*;

@Component
public class DocumentHandler {

    private static final int SPECIALTY_COLUMN = 9;
    private static final String TPL_NEXT_YEAR = "2018";
    private static final String TPL_CURRENT_YEAR = "2017";
    private static final Logger logger = LoggerFactory.getLogger(CmdRunner.class);

    private final InputStream templateAsStream = DocumentHandler.class.getResourceAsStream("/caretool_template.xlsx");

    private DateTimeFormatter monthDateFormatter = DateTimeFormatter.ofPattern("MMM yyyy", Locale.GERMANY);

    private GenerationContext context;
    private XSSFSheet tplMonthSheet;
    private XSSFSheet tplSummarySheet;
    private CreationHelper createHelper;
    private XSSFWorkbook exportWorkbook;


    public String generateDocument(GenerationContext context) throws IOException {

        this.context = context;

        //read template workbook
        exportWorkbook = new XSSFWorkbook(templateAsStream);
        createHelper = exportWorkbook.getCreationHelper();
        this.tplSummarySheet = exportWorkbook.getSheetAt(0);
        this.tplMonthSheet = exportWorkbook.getSheetAt(1);

        createTabs();

        String filename = String.format("Schichtplan %s %d.xlsx", context.getName(), context.getYear().getValue());
        exportWorkbook(filename);
        return filename;
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
                        default:
                    }
                    cell.setCellStyle(cellStyle);
                }
            }
        }
        exportWorkbook.getCTWorkbook().getSheets().getSheetArray(month.getValue()).setName(sheetName);
        logger.info("generated :" + month);
    }

    private void generateConditionalFormatting(XSSFSheet sheet) {

        XSSFSheetConditionalFormatting tplCF = tplMonthSheet.getSheetConditionalFormatting();
        XSSFSheetConditionalFormatting cf = sheet.getSheetConditionalFormatting();
        for (int i = 0; i < tplCF.getNumConditionalFormattings(); i++) {
            XSSFConditionalFormatting conditionalFormattingAt = tplCF.getConditionalFormattingAt(i);
            XSSFConditionalFormattingRule rule = conditionalFormattingAt.getRule(0);
            XSSFPatternFormatting patternFormatting = rule.getPatternFormatting();
            patternFormatting.setFillBackgroundColor(IndexedColors.YELLOW.index);
            patternFormatting.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            cf.addConditionalFormatting(conditionalFormattingAt.getFormattingRanges(), rule);
        }
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
                                    .replace(TPL_NEXT_YEAR, context.getNextYear().getValue() + "")
                                    .replace(TPL_CURRENT_YEAR, context.getCurrentYear().getValue() + "");
                            cell.setCellValue(textToAdd);
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            if (row > 7 && row != 16 && row != 12 && col > 0) {
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
        int rowNum = tplCell.getRowIndex() >= 13 ? tplCell.getRowIndex() + 25 : tplCell.getRowIndex() + 26;
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
