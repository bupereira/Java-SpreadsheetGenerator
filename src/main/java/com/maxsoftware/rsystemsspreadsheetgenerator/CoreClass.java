/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.maxsoftware.rsystemsspreadsheetgenerator;

import java.util.Calendar;
import java.util.Date;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.*;

        
/**
 *
 * @author m8416
 */
public class CoreClass {
    
    /**
     * day will store a date number (1-31) and an 8 when that's a weekday
     */
    int day[] = new int[31];
    int month, year;
    String name, supervisor, client, savePath;

    public static void main (String args[]) {

    }

    public CoreClass() {
        
    }
    
    public void evaluateDates() {
        for(int i = 1; i <= 31; i++) {
            Calendar date = Calendar.getInstance();
            date.setLenient(false); // throw an error if date is invalid
            date.setTime(new Date(year - 1900, month, i)); // -1900 because DATE adds 1900 to whatever number you give it. This is a workaround for that. https://stackoverflow.com/questions/45996752/calendar-getcalendar-year-returns-wrong-year/45996850#45996850
            if(date.get(Calendar.MONTH) != month) { // New Month
                break;
            }
            if((date.get(Calendar.DAY_OF_WEEK) > 1)&&(date.get(Calendar.DAY_OF_WEEK) < 7)) { // If Mon-Fri (Saturday is 7, Sunday is 1
                // TODO: Add holiday functionality?
                day[i - 1] = 8;
            }
        }        
    }
    
    public void setTimeVars(String month, String year) {
        this.month = Integer.parseInt(month) - 1; // For Date works with month 0-11
        this.year = Integer.parseInt(year);
    }
    
    public String getStartingValues() {
        Calendar now = Calendar.getInstance();

        Integer currMonth = now.get(Calendar.MONTH);
        Integer currYear = now.get(Calendar.YEAR);
        /* Make it human readable */
        if(currMonth == 0) { 
            currMonth = 12;
            currYear--;
        }
        
        return currMonth.toString() + "," + currYear.toString();
    }
    
    public String exportToExcel() {
        String templateFileName = "D:\\Users\\m8416\\Documents\\R Systems\\timesheet.xls";
        //Old Code ==> String newFileName = "D:\\Users\\m8416\\Documents\\R Systems\\newTimesheet.xls";
        String lastDay = getLastDay(String.valueOf(month + 1) + "/01/" + String.valueOf(year));
        savePath = savePath.replace("\\", "\\\\");
        try {
            FileInputStream excelFile = new FileInputStream(templateFileName);
            Workbook workbook = WorkbookFactory.create(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            fillCell(datatypeSheet,0,13,this.client); // 0 is row 1, 13 is col 14
            fillCell(datatypeSheet,5,2,this.name); 
            fillCell(datatypeSheet,5,13,this.supervisor); 
            fillCell(datatypeSheet,7,2,String.valueOf(month + 1) + "/01/" + String.valueOf(year)); 
            fillCell(datatypeSheet,7,14,lastDay);
            for (int i = 0; i <= 1; i++) {
                for(int j = 0; j <= 15; j++) {
                    fillCell(datatypeSheet,i + 11,j + 1,String.valueOf(day[(i * 15) + j])); 
                }
            }
            // Next 3 lines force cells with formulas to recalculate with new values
            evaluateFormulas(datatypeSheet, 11,17);
            evaluateFormulas(datatypeSheet, 12,17);
            evaluateFormulas(datatypeSheet, 18,17);
            //Old Code ==> FileOutputStream outputStream = new FileOutputStream(newFileName);
            System.out.println("Saving in :" + this.savePath);
            FileOutputStream outputStream = new FileOutputStream(this.savePath + "\\newTimeSheet.xls");
            workbook.write(outputStream);
            excelFile.close();
            outputStream.close();            
        } catch (FileNotFoundException e) {
            System.out.println("1 " + e.getMessage());
        } catch (IOException e) {
            System.out.println("2 " + e.getMessage());
        } catch (Exception e) {
            System.out.println("3 " + e.toString() + " - " + e.getLocalizedMessage() + " - " + e.getMessage());
        }

        return this.savePath;
    }

    private void fillCell(Sheet datatypeSheet, int iRow, int iCol, String text) {
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
        if(text.equals("0")) return;
        
        Row row = getNewRow(datatypeSheet, iRow);
        Cell cell = getNewCell(row, iCol);
        if(isNumeric(text)) {
            Workbook workbook = datatypeSheet.getWorkbook();
            CellStyle cs = workbook.createCellStyle();
            DataFormat df = workbook.createDataFormat();
            cs.setDataFormat(df.getFormat("0"));
            cs.setBorderBottom(cell.getCellStyle().getBorderBottom()); 
            cs.setBorderLeft(cell.getCellStyle().getBorderLeft()); 
            cs.setBorderTop(cell.getCellStyle().getBorderTop()); 
            cs.setBorderRight(cell.getCellStyle().getBorderRight()); 
            cell.setCellValue(Double.parseDouble(text)); 
            cell.setCellStyle(cs);
        } else {
            cell.setCellValue(text); 
        }
    }

    private String getLastDay(String date) {
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
        Date dt = new Date(date);
        Calendar calendar = Calendar.getInstance();  
        try {
            calendar.setTime(dt);
        }  catch(Exception e) {
            System.out.println(e.getMessage());
        }            
        calendar.add(Calendar.MONTH, 1);  
        calendar.set(Calendar.DAY_OF_MONTH, 1);  
        calendar.add(Calendar.DATE, -1);  
        Date lastDayOfMonth = calendar.getTime();  
        DateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");  
        
        return sdf.format(lastDayOfMonth);
    }

    private boolean isNumeric(String text) {
        try {  
            //System.out.println("testing..." + text);
            Integer.parseInt(text);  
            //System.out.println("True");
            return true;
        } catch(NumberFormatException e){  
            //System.out.println("False");
            return false;  
        }  
    }

    private Row getNewRow(Sheet datatypeSheet, int iRow) {
        Row row = datatypeSheet.getRow(iRow);

        return (row == null) ? datatypeSheet.createRow(iRow) : row;
    }

    private Cell getNewCell(Row row, int iCol) {
        Cell cell = row.getCell(iCol);
        
        return (cell == null) ? row.createCell(iCol) : cell;
    }

    private void evaluateFormulas(Sheet datatypeSheet, int iRow, int iCol) {
        FormulaEvaluator fe = datatypeSheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        Row row = getNewRow(datatypeSheet, iRow);
        Cell cell = getNewCell(row, iCol);
        fe.evaluateFormulaCell(cell);
    }

    public void setStringDetails(String name, String supervisor, String client, String savePath) {
        this.name = name;
        this.supervisor = supervisor;
        this.client = client;
        this.savePath = savePath;
    }
}
