package hu.webler.util;

import org.apache.poi.ss.formula.functions.Log;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelReader {

    // ez valódi konstans, az osztály tetejék van inicializálva ... láthatóság pedig jelzi, hogy kik számára érhető el
    /*
    private static final Path RESOURCES_FILE_PATH = Paths.get("src", "main", "resources");
    private static final Path EXCEL_FILE_PATH = RESOURCES_FILE_PATH.resolve("webler.xlsx");
     */
    protected static void processExcelManipulation() {
        // src/main/resources/webler.xlsx
        // ha a változó csak helyben kell (metóduson belül)
        final Path RESOURCES_FILE_PATH = Paths.get("src", "main", "resources");
        final Path EXCEL_FILE_PATH = RESOURCES_FILE_PATH.resolve("webler.xlsx");

        try {
            Workbook workbook = openExcelFile(EXCEL_FILE_PATH);
            String content = readExcelFileContent(workbook);
            Path folderPath = createFolder(RESOURCES_FILE_PATH);
            Path outputFile = folderPath.resolve("output.txt");
            writeContentToTxtFile(content, outputFile);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static Workbook openExcelFile(Path path) throws IOException {
        try (InputStream inputStream = Files.newInputStream(path)) {
            return new XSSFWorkbook(inputStream);
        }
    }

    // https://www.geeksforgeeks.org/stringbuffer-vs-stringbuilder/
    private static String readExcelFileContent(Workbook workbook) {
        StringBuilder content = new StringBuilder();
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            for (Cell cell : row) {
                content.append(cell.toString()).append("\t");
            }
            content.append("\n");
        }
        return content.toString();
    }

    // https://www.w3schools.com/java/java_date.asp
    // https://docs.oracle.com/javase/8/docs/api/java/text/DateFormat.html
    // https://www.baeldung.com/java-simple-date-format
    private static Path createFolder(Path path) throws IOException {
        String date = new SimpleDateFormat("dd-MM-yyyy").format(new Date());
        Path directoryPath = Paths.get(String.valueOf(path), date);
        if (Files.notExists(directoryPath)) {
            Files.createDirectories(directoryPath);
        }
        return directoryPath;
    }

    // https://www.geeksforgeeks.org/files-class-writestring-method-in-java-with-examples/
    // https://docs.oracle.com/en/java/javase/11/docs/api/java.base/java/nio/file/Files.html
    private static void writeContentToTxtFile(String content, Path path) throws IOException {
        Files.writeString(path, content, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);
    }
}
