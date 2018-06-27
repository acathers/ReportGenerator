import java.io.File;

import java.io.IOException;
import java.time.LocalDate;

import org.apache.poi.EncryptedDocumentException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class AncillaryMixTracker {

  public static final int MONTHLY_TRACKER_SHEET_NUMBER = 3;

  Workbook wb;
  private int startRow;
  private int startCol;
  private LocalDate currentDate = LocalDate.now();

  public AncillaryMixTracker(String filePath)
      throws EncryptedDocumentException, InvalidFormatException, IOException {

    wb = WorkbookFactory.create(new File(filePath));

    Cell cell = getCell(1, 1);
    System.out.println(cell.toString());

    if (setStartPositionForMonthlyTracker(wb)) {
      System.out.println(startRow);
      System.out.println(startCol);
    }
  }

  public Cell getCell(
      int columnNumber, int rowNumber) { // returns a cell at a specified row and column.
    int sheetNumber = 3;
    Sheet sheet = wb.getSheetAt(sheetNumber);
    Row row = sheet.getRow(rowNumber);
    Cell cell = row.getCell(columnNumber);
    return cell;
  }

  public boolean setStartPositionForMonthlyTracker(Workbook wb) {

    Sheet sheet = wb.getSheetAt(MONTHLY_TRACKER_SHEET_NUMBER);
    CellType type;

    for (Row row : sheet) {
      for (Cell cell : row) {
        type = cell.getCellTypeEnum();
        if (type == CellType.NUMERIC) {
          if (cell.getNumericCellValue() == currentDate.getDayOfMonth()) {
            startRow = cell.getRowIndex();
            startCol = cell.getColumnIndex();
            return true;
          }
        }
      }
    }
    return false;
  }
}
