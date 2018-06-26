package org.build.ReportGenerator;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;

import org.apache.poi.EncryptedDocumentException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class OperationsReport {
  Workbook wb;

  private static final int VEHICLE_COUNT_OFFSET =
      5; // these values are offsets from the cell containing "For specified date range"
  private static final int NET_SALES_OFFSET = 9;
  private static final int LABOR_HOURS_OFFSET = 1;
  private static final int LABOR_DOLLARS_OFFSET = 2;

  private Cell netSalesCell;
  private Cell vehicleCountCell;
  private Cell laborHoursCell;
  private Cell laborDollarsCell;

  private double netSales;
  private double vehicleCount;
  private double laborPerVehicle;
  private double ticketAverage;
  private double laborNetSales;
  private double laborHours;
  private double laborDollars;

  private int startRow; // starting row number for cell containing "for specified date range"
  private int startCol; // starting column number for cell containing "For specified date range"

  public OperationsReport(String filePath)
      throws InvalidFormatException, EncryptedDocumentException, IOException {
    wb = WorkbookFactory.create(new File(filePath));

    extractReportValues(wb); // extracts minimum amount of data to do furthur calculations
    setReportValues();
  }

  public void extractReportValues(
      Workbook wb) { // extracts mimimum amount of data to do furthur calculations.

    if (setStartRowAndCol(wb)) {
      netSalesCell = getCell(startCol, startRow + NET_SALES_OFFSET);
      vehicleCountCell = getCell(startCol, startRow + VEHICLE_COUNT_OFFSET);
      laborHoursCell = getCell(startCol, startRow + LABOR_HOURS_OFFSET);
      laborDollarsCell = getCell(startCol, startRow + LABOR_DOLLARS_OFFSET);

    } else {
      System.out.print("Unable to run, couldn't find required cell to generate values");
    }
  }

  public void setReportValues() { // sets all values required by either grabbing directly from excel
    // document, or doing calculation
    netSales = getCellValue(netSalesCell);
    vehicleCount = getCellValue(vehicleCountCell);
    laborHours = getCellValue(laborHoursCell);
    laborDollars = getCellValue(laborDollarsCell);
    ticketAverage = calculateTicketAverage();
    laborPerVehicle = calculateLaborPerVehicle();
    laborNetSales = calculateLaborNetSales();
  }

  public double getNetSales() {
    return netSales;
  }

  public double getVehicleCount() {
    return vehicleCount;
  }

  public double getLaborPerVehicle() {
    return laborPerVehicle;
  }

  public double getTicketAverage() {
    return ticketAverage;
  }

  public double getLaborNetSales() {
    return laborNetSales;
  }

  public double getLaborHours() {
    return laborHours;
  }

  public double getLaborDollars() {
    return laborDollars;
  }

  public double calculateTicketAverage() {

    return netSales / vehicleCount;
  }

  private double calculateLaborNetSales() {
    return laborDollars / netSales;
  }

  public double calculateLaborPerVehicle() { // returns value to 2 decimal places
    double value = laborHours / vehicleCount;
    BigDecimal bd = new BigDecimal(value).setScale(2, RoundingMode.HALF_EVEN);
    return bd.doubleValue();
  }

  public double getCellValue(Cell cell) { // returns double value for a cell.
    return cell.getNumericCellValue();
  }

  public Cell getCell(
      int columnNumber, int rowNumber) { // returns a cell at a specified row and column.
    int sheetNumber = 0;
    Sheet sheet = wb.getSheetAt(sheetNumber);
    Row row = sheet.getRow(rowNumber);
    Cell cell = row.getCell(columnNumber);
    return cell;
  }

  public boolean setStartRowAndCol(
      Workbook wb) { // sets values for startRow and startCol and returns a boolean if cell is found
    Sheet sheet = wb.getSheetAt(0);
    DataFormatter dataFormatter = new DataFormatter();

    for (Row row : sheet) {
      for (Cell cell : row) {
        if (dataFormatter.formatCellValue(cell).equals("For Specified Date Range")) {
          startRow = cell.getRowIndex();
          startCol = cell.getColumnIndex();
          return true;
        }
      }
    }
    return false;
  }

  public Cell getNetSalesCell() {
    return netSalesCell;
  }

  public Cell getVehicleCountCell() {
    return vehicleCountCell;
  }

  public Cell getLaborDollarsCell() {
    return laborDollarsCell;
  }
}
