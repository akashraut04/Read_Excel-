import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelDuplicateFinder {

    public static void main(String[] args) {
        String filePath = "C:\\Users\\Akash.Raut\\OneDrive - NEC Software Solutions\\Documents\\New Folder\\ASH3.xlsx";
        String columnName = "SPARGO Company Name";

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            Map<String, Integer> valueCounts = new HashMap<>();

            Row headerRow = sheet.getRow(0);
            int columnIndex = -1;
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                    columnIndex = cell.getColumnIndex();
                    System.out.println("Column: " + columnIndex);
                    break;
                }
            }

            if (columnIndex == -1) {
                System.out.println("Column not found: " + columnName);
                return;
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null) {
                        String cellValue = cell.toString().trim();
                        valueCounts.put(cellValue, valueCounts.getOrDefault(cellValue, 0) + 1);
                    }
                }
            }

            System.out.println("All Records with Counts:");
            valueCounts.forEach((key, count) ->
                    System.out.println("Value: " + key + " | Count: " + count)
            );

            List<Map.Entry<String, Integer>> top100Records = valueCounts.entrySet()
                    .stream()
                    .sorted((a, b) -> b.getValue().compareTo(a.getValue()))
                    .limit(100)
                    .toList();

            System.out.println("\nTop 100 Most Frequent Values:");
            for (Map.Entry<String, Integer> entry : top100Records) {
                System.out.println("Value: " + entry.getKey() + " | Count: " + entry.getValue());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
