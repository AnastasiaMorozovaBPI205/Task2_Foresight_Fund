import com.gargoylesoftware.htmlunit.WebClient;
import com.gargoylesoftware.htmlunit.html.HtmlAnchor;
import com.gargoylesoftware.htmlunit.html.HtmlElement;
import com.gargoylesoftware.htmlunit.html.HtmlPage;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        getFiles();
        parseFiles();
    }

    static void getFiles() throws IOException {
        WebClient client = new WebClient();
        client.getOptions().setCssEnabled(false);
        client.getOptions().setJavaScriptEnabled(false);

        int count = 0;
        for (int i = 1; i < 5; ++i) {
            HtmlPage page = client.getPage("https://bashesk.ru/corporate/tariffs/unregulated/?PAGEN_1="+ i +"&filter_name=&filter_date_from=01.07.2019&filter_date_to=01.06.2020");

            List<HtmlElement> items = page.getByXPath("//div[@class='col-2']");

            for (HtmlElement item : items) {
                HtmlAnchor itemAnchor = item.getFirstByXPath(".//a");
                String itemUrl = itemAnchor.getHrefAttribute();

                if (itemUrl.contains("ПУНЦЭМ_до 670кВт")) {
                    System.out.println(itemUrl);
                    try (InputStream in = new URL("https://bashesk.ru/" + itemUrl.replace(" ", "%20")).openStream()) {
                        Files.copy(in, Paths.get(count + ".xls"));
                        ++count;
                    }
                }
            }
        }
    }

    static void parseFiles() throws IOException {
        for (int i = 0; i < 13; ++i) {
            HSSFWorkbook excelBook = new HSSFWorkbook(new FileInputStream(i + ".xls"));
            HSSFSheet bookSheet = excelBook.getSheetAt(0);
            boolean cellIsFound = false;

            for (int j = 0; j < bookSheet.getLastRowNum(); ++j) {
                HSSFRow row = bookSheet.getRow(j);

                for (int k = 0; k < row.getLastCellNum(); ++k) {
                    if (row.getCell(k) == null) {
                        continue;
                    }

                    if (row.getCell(k).getCellType() == 1 && row.getCell(k).getStringCellValue().contains("г) объем фактического пикового потребления гарантирующего поставщика" +
                            " на оптовом рынке, МВт")) {
                        cellIsFound = true;

                        System.out.println(row.getCell(k + 15).getNumericCellValue());
                        break;
                    }
                }

                if (cellIsFound) {
                    break;
                }
            }
        }
    }
}
