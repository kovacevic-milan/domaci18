import com.github.javafaker.Faker;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class Main {
    public static void main(String[] args) {

        String relativePath = "domaci18.xlsx";
        try {
            citanjeFajla(relativePath);
            ispisFajla(relativePath);
            dodajOsobu(relativePath);
            iscitavanjeOsoba(relativePath);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void citanjeFajla(String relativePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(relativePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("imena");

        for (int i = 0; i < 5; i++) {
            XSSFRow row1 = sheet.getRow(i);
            for (int j = 0; j < 2; j++) {
                XSSFCell cell1 = row1.getCell(j);
                System.out.print(cell1.getStringCellValue() + " ");
            }
            System.out.println();
        }
    }

    public static void ispisFajla(String filename) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("imena");
        for (int i = 0; i < 5; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.getCell(0);
            XSSFCell cell1 = row.getCell(1);
            XSSFCell cell2 = row.createCell(2);
            XSSFCell cell3 = row.createCell(3);
            cell2.setCellValue(String.valueOf(cell));
            cell3.setCellValue(String.valueOf(cell1));
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filename);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public static void dodajOsobu(String filename) throws IOException {
        Faker faker = new Faker();
        String name1 = faker.name().firstName();
        String name2 = faker.name().firstName();
        String name3 = faker.name().firstName();
        String name4 = faker.name().firstName();
        String name5 = faker.name().firstName();

        String lastName1 = faker.name().lastName();
        String lastName2 = faker.name().lastName();
        String lastName3 = faker.name().lastName();
        String lastName4 = faker.name().lastName();
        String lastName5 = faker.name().lastName();

        ArrayList<String> ime = new ArrayList<>();
        ime.add(name1);
        ime.add(name2);
        ime.add(name3);
        ime.add(name4);
        ime.add(name5);

        ArrayList<String> prezime = new ArrayList<>();
        prezime.add(lastName1);
        prezime.add(lastName2);
        prezime.add(lastName3);
        prezime.add(lastName4);
        prezime.add(lastName5);

        FileInputStream fileInputStream = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("imena");
        for (int i = 0; i < 5; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell2 = row.createCell(4);
            XSSFCell cell3 = row.createCell(5);
            cell2.setCellValue(ime.get(i));
            cell3.setCellValue(prezime.get(i));
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filename);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public static void iscitavanjeOsoba(String relativePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(relativePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("imena");

        for (int i = 0; i < 5; i++) {
            XSSFRow row1 = sheet.getRow(i);
            for (int j = 4; j < 6; j++) {
                XSSFCell cell1 = row1.getCell(j);
                System.out.print(cell1.getStringCellValue() + " ");
            }
            System.out.println();
        }
    }
}



