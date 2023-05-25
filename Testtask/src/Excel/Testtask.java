package Excel;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Testtask {

	public static void main(String[] args) {

        try (Workbook workbook = new XSSFWorkbook()) {
        
            Sheet sheet = workbook.createSheet("Sheet1");

            for (int i = 0; i < 256; i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < 26; j++) {
                    Cell cell = row.createCell(j);
                    if ((i + j) % 2 == 0) {
                        int intValue = i * j;
                        cell.setCellValue(intValue);
                    } else {
                        double doubleValue = Math.pow(i, j);
                        cell.setCellValue(doubleValue);
                    }

                    if (i == 0) {
  
                        String columnName = Character.toString((char) ('A' + j));
                        cell.setCellValue(columnName);
                    } else if (j == 0) {
                        cell.setCellValue(i);
                    } else {
                        String formula = "A" + (i + 1) + "*" + (j + 1);
                        cell.setCellFormula(formula);
                    }
                }
            }

            CellStyle style = workbook.createCellStyle();
            style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
            for (int i = 0; i < 256; i++) {
                Row row = sheet.getRow(i);
                for (int j = 0; j < 26; j++) {
                    Cell cell = row.getCell(j);
                    if ((i + j) % 2 != 0) {
                        cell.setCellStyle(style);
                    }
                }
            }

            for (int i = 0; i < 26; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream outputStream = new FileOutputStream("workbook.xlsx")) {
                workbook.write(outputStream);
            } 
            catch (Exception e) {
                e.printStackTrace();
            }
        } catch (IOException e1) {
			e1.printStackTrace();
		}
    }
}