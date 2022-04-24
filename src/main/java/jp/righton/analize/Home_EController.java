package jp.righton.analize;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
@Controller
public class Home_EController {
    //　　　　　　　　　　　　　全社順位、　　　ブロック順位、　　　　自店順位、　　　部門、　
    record TaskItem_E(Cell rankCompany, Cell rankBlock, Cell rankStore, Cell Group,
                      //     品番、          品名、        当週売点、    店舗在庫           商品画像リンク
                      Cell itemNumber, Cell itemName, Cell salesPoint, Cell stock, String itemLink){}
    public List<Home_EController.TaskItem_E> taskItems1 = new ArrayList<>(); //売れ筋
    public List<Home_EController.TaskItem_E> taskItems2 = new ArrayList<>(); //売れ筋候補
    public List<Home_EController.TaskItem_E> taskItems3 = new ArrayList<>(); //店舗特性

    @RequestMapping("/list_E")
    String listItems_E(Model model) {
        model.addAttribute("taskList_E1", taskItems1); //売れ筋
        model.addAttribute("taskList_E2", taskItems2); //売れ筋候補
        model.addAttribute("taskList_E3", taskItems3); //店舗特性
        return "home_E";
    }
    @PostMapping("/add_E")
    public String analyze(@RequestPart("UpLoadFile_E") MultipartFile UpLoadFile)throws EncryptedDocumentException, IOException {
        //ファイル名取得
        String fileName = UpLoadFile.getOriginalFilename();
        //デスクトップのパス取得　　全社システムにするならデスクトップパスを店舗PCに合うものに変更⇒[D:\\店舗用\\Desktop\\]
        Path dtPath = Paths.get("C:\\Users\\yoshi\\Desktop");
        //エクセルファイルへアクセス⇒デスクトップのパス+ファイル名　※" / " これ漏れたらNotFileFound！
        Workbook excel = WorkbookFactory.create(new File(dtPath + "/" + fileName ));
        // エクセルシート名
        Sheet sheet = excel.getSheet("店別");
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
            //29番目のセル（店舗在庫）
            Cell stock = row.getCell(29);

            double numRankCompany = Double.parseDouble(String.valueOf(rankCompany));
            double numRankStore = Double.parseDouble(String.valueOf(rankStore));
            double numSalesPoint = Double.parseDouble(String.valueOf(salesPoint));
            double numRankBlock = Double.parseDouble(String.valueOf(rankBlock));

            //売れ筋条件 全社順位3位以上かつ自店順位3位以上かつ当週売れ点1以上
            if (numSalesPoint >= 1.0 && numRankCompany <= 3.0 && numRankStore <= 3.0) {
                String itemLink = "https://right-on.co.jp/search?q=" + itemNumber;
                Home_EController.TaskItem_E item1 = new Home_EController.TaskItem_E(rankCompany, rankBlock, rankStore,
                        Group, itemNumber, itemName, salesPoint, stock, itemLink);
                taskItems1.add(item1);
            }
            //売れ筋候補条件 全社順位3位以上かつブロック順位3位以上かつ自店順位10位以下
            if (numRankCompany <= 5.0 && numRankBlock <= 5.0 && numRankStore >= 7.0) {
                String itemLink = "https://right-on.co.jp/search?q=" + itemNumber;
                Home_EController.TaskItem_E item2 = new Home_EController.TaskItem_E(rankCompany, rankBlock, rankStore,
                        Group, itemNumber, itemName, salesPoint, stock, itemLink);
                taskItems2.add(item2);
            }
            //店舗特性条件 全社順位10位以下かつブロック順位10位以下かつ当週売れ点8点以上かつ自店順位3位以上かつ当週売れ点1以上
            if (numRankCompany >= 10.0 && numRankBlock >= 10.0 && numSalesPoint >= 1.0 && numRankStore <= 3.0) {
                String itemLink = "https://right-on.co.jp/search?q=" + itemNumber;
                Home_EController.TaskItem_E item3 = new Home_EController.TaskItem_E(rankCompany, rankBlock, rankStore,
                        Group, itemNumber, itemName, salesPoint, stock, itemLink);
                taskItems3.add(item3);
            }
        }return "redirect:/list_E";
    }
}