package org.build.ReportGenerator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.temporal.ChronoField;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;

public class ReportGenerator {
  private static final String OPERATIONS_REPORT_FILE_PATH = "Reports/OperationsReport.xls";
  private static final String EMAIL_STATS_FILE_PATH = "EmailStats.xls";

  private final LocalTime currentTime = LocalTime.now();
  private final LocalTime sixPM = LocalTime.of(18, 00);
  private final LocalTime fivePM = LocalTime.of(17, 00);
  private final LocalTime fourPM = LocalTime.of(16, 00);
  private final LocalTime twoPM = LocalTime.of(14, 00);
  private final LocalTime noon = LocalTime.of(12, 00);
  private final int vehicleCount = 1;
  private final int netSales = 2;
  private final int ticketAverage = 3;
  private final int laborPerVehicle = 5;
  private final int laborNetSales = 8;

  public void run() {
    try (FileInputStream emailStats = new FileInputStream(new File(EMAIL_STATS_FILE_PATH))) {

      updateHourlyTracker(emailStats);

    } catch (EncryptedDocumentException e) {
      System.out.println("File is encrypted");
    } catch (InvalidFormatException e) {
      System.out.println("Invalid file type, xls only");
    } catch (IOException e) {
      System.out.println("file not found or is in use");
      e.printStackTrace();
    }
  }

  public void updateHourlyTracker(FileInputStream emailStats)
      throws EncryptedDocumentException, InvalidFormatException, IOException {

    OperationsReport report = new OperationsReport(OPERATIONS_REPORT_FILE_PATH);

    HSSFWorkbook wb = new HSSFWorkbook(emailStats);
    HSSFSheet worksheet = wb.getSheetAt(0);
    Cell cell = null;

    updateAllRowValues(report, worksheet, cell);
    emailStats.close();
    writeToWorkbook(wb, new FileOutputStream(new File(EMAIL_STATS_FILE_PATH)));
  }

  public void updateCell(Cell cell, double value) {
    cell.setCellValue(value);
  }

  public void updateAllRowValues(OperationsReport report, HSSFSheet worksheet, Cell cell) {

    if (getRowForDayAndTime() != -1) {
      cell = worksheet.getRow(getRowForDayAndTime()).getCell(vehicleCount);
      updateCell(cell, report.getVehicleCount());
      cell = worksheet.getRow(getRowForDayAndTime()).getCell(netSales);
      updateCell(cell, report.getNetSales());
      cell = worksheet.getRow(getRowForDayAndTime()).getCell(ticketAverage);
      updateCell(cell, report.getTicketAverage());
      cell = worksheet.getRow(getRowForDayAndTime()).getCell(laborPerVehicle);
      updateCell(cell, report.getLaborPerVehicle());
      cell = worksheet.getRow(getRowForDayAndTime()).getCell(laborNetSales);
      updateCell(cell, report.getLaborNetSales());
    } else {
      System.out.println("Too early to run report");
    }
  }

  public void writeToWorkbook(HSSFWorkbook wb, FileOutputStream outputFile) throws IOException {
    wb.write(outputFile);
    outputFile.close();
  }

  public int getRowForDayAndTime() {
    DayOfWeek dayOfWeek = DayOfWeek.of(LocalDate.now().get(ChronoField.DAY_OF_WEEK));

    if (dayOfWeek.equals(DayOfWeek.SATURDAY) && currentTime.isAfter(fivePM)) {
      return 13;
    } else if (dayOfWeek.equals(DayOfWeek.SUNDAY) && currentTime.isAfter(fourPM)) {
      return 13;
    } else if (currentTime.isAfter(sixPM)) {
      return 13;
    } else if (currentTime.isAfter(fourPM)) {
      return 9;
    } else if (currentTime.isAfter(twoPM)) {
      return 7;
    } else if (currentTime.isAfter(noon)) {
      return 5;
    } else {
      return -1;
    }
  }
}
