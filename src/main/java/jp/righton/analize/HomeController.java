package jp.righton.analize;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

//プロJavaではHomeコントローラーの（”/hello")
@Controller
public class HomeController {
    record TaskItem(Cell rankCompany, Cell rankBlock, Cell rankStore, Cell Group,
                    Cell itemNumber, Cell itemName, Cell salesPoint){}
    public List<TaskItem> taskItems1 = new ArrayList<>();
    public List<TaskItem> taskItems2 = new ArrayList<>();
    public List<TaskItem> taskItems3 = new ArrayList<>();


    @RequestMapping("/list")
    String listItems(Model model) {
                model.addAttribute("taskList1", taskItems1);
                model.addAttribute("taskList2", taskItems2);
                model.addAttribute("taskList3", taskItems3);
                return "home";
    }


    @PostMapping("/add")
    public String bestSeller(@RequestParam("UpLoadFile") MultipartFile UpLoadFile)throws EncryptedDocumentException, IOException {
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

            double numRankCompany = Double.parseDouble(String.valueOf(rankCompany));
            double numRankStore = Double.parseDouble(String.valueOf(rankStore));
            double numSalesPoint = Double.parseDouble(String.valueOf(salesPoint));
            double numRankBlock = Double.parseDouble(String.valueOf(rankBlock));

            //売れ筋条件
            if (numSalesPoint >= 5.0 && numRankCompany <= 3 && numRankStore <= 3.0) {
                TaskItem item1 = new TaskItem(rankCompany, rankBlock, rankStore,
                        Group, itemNumber, itemName, salesPoint);
                taskItems1.add(item1);
            }
            //売れ筋候補条件
            if (numRankCompany <= 3.0 && numRankBlock <= 3.0 && numRankStore >= 10.0) {
                TaskItem item2 = new TaskItem(rankCompany, rankBlock, rankStore,
                        Group, itemNumber, itemName, salesPoint);
                taskItems2.add(item2);
            }
            //店舗特性条件
            if (numRankBlock <= 3.0 && numRankCompany >= 6.0 && numRankStore <= 3.0) {
                TaskItem item3 = new TaskItem(rankCompany, rankBlock, rankStore,
                        Group, itemNumber, itemName, salesPoint);
                taskItems3.add(item3);
        }}
        return "redirect:/list";
    }
   /* String addItem( @RequestParam("UpLoadFile") MultipartFile multipartFile ,
                   @RequestParam("rankCompany") String numRankCompany,
                   @RequestParam("numRankBlock") String numRankBlock,
                   @RequestParam("numRankStore") String numRankStore,
                   @RequestParam("Group") String Group,
                   @RequestParam("itemNumber") String itemNumber,
                   @RequestParam("itemName") String itemName,
                   @RequestParam("numSalesPoint") double numSalesPoint){
        String id = UUID.randomUUID().toString().substring(0, 8);
        TaskItem item = new TaskItem(id, numRankCompany, numRankBlock, numRankStore,
                Group, itemNumber, itemName, numSalesPoint);
        taskItems.add(item);*/


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

