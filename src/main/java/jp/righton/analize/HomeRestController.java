package jp.righton.analize;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.bind.annotation.RequestMapping;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.stream.Collectors;
import org.springframework.web.bind.annotation.PostMapping;
@RestController
public class HomeRestController {

    record TaskItem(String id, String numRankCompany, String numRankBlock, String numRankStore, String Group,
                    String itemNumber, String itemName, double numSalesPoint){}
    private List<TaskItem> taskItems = new ArrayList<>();


    @RequestMapping(value = "/resthello")
         String hello() {
        return """
                Hello.
                It works!
                現在時刻は%sです。
                """.formatted(LocalDateTime.now());
    }

        @PostMapping("/restadd")
        String addItem ( @RequestParam("numRankCompany") String numRankCompany,
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

            return "データを追加しました。";
        }

        @PostMapping("/restlist")
        String ListItems () {
            String result = taskItems.stream()
                    .map(TaskItem::toString)
                    .collect(Collectors.joining(", "));
            return "result";
        }
        String hello1 () {
            return """
                    Hello.
                    It works!
                    現在時刻は%sです。
                    """.formatted(LocalDateTime.now());
        }
        /*@RequestMapping("/list /add /home")
         String ListItems () {
            String result = analyzeInfo.stream()
                    .map(HomeRestController.AnalyzeInfo::toString)
                    .collect(Collectors.joining(", "));
            return result;
        }*/


    }
