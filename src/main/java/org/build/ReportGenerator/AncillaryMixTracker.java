package org.build.ReportGenerator;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class AncillaryMixTracker {

  Workbook wb;
  private int startRow;
  private int startCol;

  public AncillaryMixTracker(String filePath)
      throws EncryptedDocumentException, InvalidFormatException, IOException {

    wb = WorkbookFactory.create(new File(filePath));
    if(setStartPositionForMonthlyTracker(wb)) {
        System.out.println(startRow);
        System.out.println(startCol);
    }
  }

  public boolean setStartPositionForMonthlyTracker(
      Workbook wb) { // sets values for startRow and startCol and returns a boolean if cell is found
    LocalDate currentDate = LocalDate.now();
    Sheet sheet = wb.getSheetAt(2);
    DataFormatter dataFormatter = new DataFormatter();

    for (Row row : sheet) {
      for (Cell cell : row) {
        if (dataFormatter.formatCellValue(cell).equals(currentDate.getDayOfMonth())) {
          startRow = cell.getRowIndex();
          startCol = cell.getColumnIndex();
          return true;
        }
      }
    }
    return false;
  }
}
