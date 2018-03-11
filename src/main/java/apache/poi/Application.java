package apache.poi;

import apache.poi.reader.ExcelReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;

public class Application {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        String path = "/home/bartek/Desktop/RegonTest.xlsx";
        String sheetName = "Regon";
        ExcelReader reader = new ExcelReader(path,sheetName);
        int regonCell = reader.getCellIndexWithText("Regon");
        int[] ints = {1,2,3,4,5};
        reader.writeEmptyRow("Inty", ints);
    }
}
