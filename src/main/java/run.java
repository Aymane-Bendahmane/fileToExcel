import java.util.Scanner;
import java.util.logging.Logger;

public class run {
    private static Logger logger = Logger.getLogger("Logging");

    public static void main(String[] args) throws Exception {
        logger.info("Excel Name");
        Scanner scanner = new Scanner(System.in);
        FileReaderToExcel fileReaderToExcel = new FileReaderToExcel(scanner.next());
        logger.info("Reading the input file : ");
        fileReaderToExcel.readFile();
        logger.warning("loooooool");
    }
}
