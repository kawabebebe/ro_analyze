package jp.righton.analize;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.io.File;
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
    public String analyze(@RequestPart("UpLoadFile") MultipartFile UpLoadFile)throws EncryptedDocumentException, IOException {
        //Cドライブ直下にファイルを保存

        //Cドライブ直下のパス取得
        Path savePath = Paths.get("C:/");
        //ファイル名取得
        String fileName = UpLoadFile.getOriginalFilename();
        //エクセルファイルへアクセス(Cドライブ直下のパス名+ファイル名)
        Workbook excel = WorkbookFactory.create(new File(savePath + fileName ));
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

            //売れ筋条件 全社順位3位以上かつ自店順位3位以上かつ当週売点5点以上
            if (numSalesPoint >= 5.0 && numRankCompany <= 3 && numRankStore <= 3.0) {
                TaskItem item1 = new TaskItem(rankCompany, rankBlock, rankStore,
                        Group, itemNumber, itemName, salesPoint);
                taskItems1.add(item1);
            }
            //売れ筋候補条件 全社順位3位以上かつブロック順位3位以上かつ自店順位10位以下
            if (numRankCompany <= 3.0 && numRankBlock <= 3.0 && numRankStore >= 10.0) {
                TaskItem item2 = new TaskItem(rankCompany, rankBlock, rankStore,
                        Group, itemNumber, itemName, salesPoint);
                taskItems2.add(item2);
            }
            //店舗特性条件 全社順位10位以下かつブロック順位10位以下かつ当週売れ点3点以上かつ自店順位3位以上
            if (numRankCompany >= 10.0 && numRankBlock >= 10.0 && numSalesPoint >= 3.0 && numRankStore <= 3.0) {
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

