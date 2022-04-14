package jp.righton.analize;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.bind.annotation.RequestMapping;
import javax.swing.*;
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

