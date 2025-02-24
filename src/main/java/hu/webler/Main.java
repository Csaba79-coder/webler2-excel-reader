package hu.webler;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Main {

    // first step add dependency to pom.xml
    // dependency 1: org.apache.poi - poi
    // dependency 2: org.apache.poi - poi-ooxml
    // https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/usermodel/XSSFWorkbook.html
    // https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Workbook.html
    public static void main(String[] args) {

        String filePath = "src/main/resources/webler.xlsx";

        File file = new File(filePath);

        try (FileInputStream inputStream = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0); // Assuming that we need the first sheet!
            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println(); // move to the next row
            }
        } catch (IOException e) {
            System.out.println("Error reading excel file: " + e.getMessage());
        }
    }
}
