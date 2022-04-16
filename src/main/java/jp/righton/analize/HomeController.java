package jp.righton.analize;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.bind.annotation.RequestMapping;
import javax.swing.*;
import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

//プロJavaではHomeコントローラーの（”/hello")
@Controller
public class HomeController {
    record TaskItem(String id, String numRankCompany, String numRankBlock, String numRankStore, String Group,
                    String itemNumber, String itemName, double numSalesPoint){}
    private List<TaskItem> taskItems = new ArrayList<>();

    @RequestMapping("/list")
    String listItems(Model model) {
                model.addAttribute("taskLists", taskItems);
                return "home";
    }

    @GetMapping("/add")
    public static String bestSeller(@RequestParam("UpLoadFile")String UpLoadFile)throws EncryptedDocumentException, IOException {
        //エクセルファイルへアクセス
        Workbook excel = WorkbookFactory.create(new File("C://tenpo_hinban_jisseki_1319_202230.xlsx"));
        // エクセルシート名
        Sheet sheet = excel.getSheet("店別");

        //~~~~ここから~~~~//
        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
            //6行目(セル順は５)以降が必要データ
            Row row = sheet.getRow(i);
            //4番目のセル（全社順位）
            Cell rankCompany = row.getCell(3);
            //5番目のセル（ブロック順位）
            Cell rankBlock = row.getCell(4);
            //6番目のセル（店舗順位）
            Cell rankStore = row.getCell(5);
            //12番目のセル（部門）
            Cell Group = row.getCell(11);
            //18番目のセル(品番)
            Cell itemNumber = row.getCell(17);
            //19番目のセル（品名）
            Cell itemName = row.getCell(18);
            //24番目のセル（当週売点）
            Cell salesPoint = row.getCell(23);


            //全社順位、ブロック順位、自店順位,当週売れ点を実数（.いくつ表記）に変換

            double numRankCompany = Double.parseDouble(String.valueOf(rankCompany));
            double numRankBlock = Double.parseDouble(String.valueOf(rankBlock));
            double numRankStore = Double.parseDouble(String.valueOf(rankStore));
            double numSalesPoint = Double.parseDouble(String.valueOf(salesPoint));

            if (numSalesPoint >= 5.0 && numRankCompany <= 3 && numRankStore <= 3.0) {
                //取得した文字列の表示
                System.out.println("売れ筋アイテム");
                System.out.println("全社順位 : " + numRankCompany);
                System.out.println("ブロック順位 : " + numRankBlock);
                System.out.println("自店順位 : " + numRankStore);
                System.out.println("部門 : " + Group);
                System.out.println("品番 : " + itemNumber);
                System.out.println("品名 : " + itemName);
                System.out.println("当週売点 : " + numSalesPoint);
                System.out.println("----------------------");
            }
        }

        return UpLoadFile;
    }


    String addItem( /*@RequestParam("UpLoadFile") MultipartFile multipartFile , */
                   @RequestParam("numRankCompany") String numRankCompany,
                   @RequestParam("numRankBlock") String numRankBlock,
                   @RequestParam("numRankStore") String numRankStore,
                   @RequestParam("Group") String Group,
                   @RequestParam("itemNumber") String itemNumber,
                   @RequestParam("itemName") String itemName,
                   @RequestParam("numSalesPoint") double numSalesPoint){
        String id = UUID.randomUUID().toString().substring(0, 8);
        TaskItem item = new TaskItem(id, numRankCompany, numRankBlock, numRankStore,
                Group, itemNumber, itemName, numSalesPoint);
        taskItems.add(item);

        return "redirect:/list";
    }
/*
    @RequestMapping("/list")
    String ListItems(){
        String result = analyzeInfo.stream()
               .map(AnalyzeInfo::toString)
                .collect(Collectors.joining(", "));
       return result;
    }
*/


    @RequestMapping(value="/hello")
    String hello(Model model) {
        model.addAttribute("time",LocalDateTime.now());
        return "hello";
    }


}

