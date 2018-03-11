package apache.poi.reader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {

    private Workbook excel;
    private Sheet sheet;

    public ExcelReader(String path, String sheetName) throws IOException, InvalidFormatException {
        this.excel = WorkbookFactory.create(new FileInputStream(new File(path)));
        this.sheet = excel.getSheet(sheetName);
    }

    public int getCellIndexWithText(String param) {
      Row firstRow = sheet.getRow(1);
      List<Integer> integerList = new ArrayList<>();
      int lastCollumn = firstRow.getLastCellNum();
        for (int i = 0; i < lastCollumn; i++) {
            Cell cell = firstRow.getCell(i);
            if (cell.getStringCellValue().equals(param)) {
                integerList.add(i);
            }
        }
        return integerList.stream()
                .findFirst()
                .orElseThrow(
                        () -> new RuntimeException("Brak Tablicy o wartości " + param ));
    }

    public List<Double> getRegonValueFormExcel(int cellNo) {
      int lastRow = sheet.getLastRowNum();
      List<Double> regonList = new ArrayList<>();
        for (int i = 2; i < lastRow + 1 ; i++) {
           regonList.add(sheet.getRow(i).getCell(cellNo).getNumericCellValue());
        }
        return regonList;
    }

    public void writeEmptyRow(String cellName, int... param) {
      Row firstRow = sheet.getRow(1);
      int writtenCellNumber = firstRow.getLastCellNum() +1;
      Cell writtenCell = firstRow.createCell(writtenCellNumber, CellType.STRING);
      writtenCell.setCellValue(cellName);
        for (int i = 2; i < param.length +2 ; i++) {
            sheet.getRow(i)
                 .createCell(writtenCellNumber, CellType.NUMERIC)
                 .setCellValue(param[i-2]);
        }
    }


}