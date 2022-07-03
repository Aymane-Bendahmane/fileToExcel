import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class FileReaderToExcel {
    private String fileTxt;
    String excelFilePath ;
    FileInputStream inputStream;
    Workbook workbook ;

    public FileReaderToExcel(String fileName) throws IOException {
        excelFilePath = fileName;
        inputStream  = new FileInputStream(excelFilePath);
        workbook = WorkbookFactory.create(inputStream);
    }

    public void readFile() throws Exception {


        List<String> data = new ArrayList<>();
        System.out.println("Please enter the text name : ");
        Scanner scanner = new Scanner(System.in);

        this.fileTxt = scanner.next();

        File file = new File(fileTxt);
        BufferedReader br
                = new BufferedReader(new FileReader(file));
        String st = br.readLine();
        while (st != null && !st.toUpperCase(Locale.ROOT).contains("**SPLIT")) {
            if (!st.isEmpty()) {
                data.add(st);
            }
            st = br.readLine();
        }
        this.dataToExcel(data);
        data.clear();
        inputStream.close();

        FileOutputStream outputStream = new FileOutputStream(excelFilePath);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }


    public void dataToExcel(List<String> data) {

        try {
            Sheet sheet = workbook.getSheetAt(0);
            Row newRow = sheet.createRow(1);
            for (int i = 0; i < 19; i++) {
                Cell cell = newRow.createCell(i);
                switch (i) {
                    case 4://fisrtNAme
                        cell = newRow.createCell(4);
                        cell.setCellValue(data.get(0));
                        break;
                    case 5://fisrtNAme
                        cell = newRow.createCell(5);
                        cell.setCellValue(data.get(0));
                        break;
                    case 6://Addresse
                        cell = newRow.createCell(6);
                        cell.setCellValue(data.get(1));
                        break;
                    case 10://phone
                        cell = newRow.createCell(10);
                        data.set(2, this.checkFormatPhone(data.get(2)));
                        cell.setCellValue(data.get(2));
                        break;
                    default:
                        cell.setCellValue("");
                        break;
                }
            }
        } catch (EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }

    String checkFormatPhone(String s) {
        String prefix = "+212";
        if (s == null || s.isEmpty())
            return "";
        String toCheck = s.substring(0,4);
        if (!prefix.equals(toCheck))
            return prefix.concat(s.substring(1));
        return "";
    }

}
