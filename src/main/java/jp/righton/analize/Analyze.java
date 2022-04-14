package jp.righton.analize;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Analyze {

    /*public static void main(String[] args) throws EncryptedDocumentException, IOException {*/

    public static void bestSeller()throws EncryptedDocumentException, IOException {
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




            /*
            record ItemInfo(double numRankCompany, double numRankBlock, double numRankStore, String Group,
                                        String itemNumber, String itemName, double numSalesPoint){}
            new itemInfo()
            */


            //全社順位3位以上かつ自店順位3位以上かつ当週売点5点以上（売れ筋）

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

            //全社順位3位以上かつブロック順位3位以上かつ自店順位10位以下（売れ筋候補）
            if (numRankCompany <= 3.0 && numRankBlock <= 3.0 && numRankStore >= 10.0) {
                //取得した文字列の表示
                System.out.println("売れ筋候補アイテム");
                System.out.println("全社順位 : " + numRankCompany);
                System.out.println("ブロック順位 : " + numRankBlock);
                System.out.println("自店順位 : " + numRankStore);
                System.out.println("部門 : " + Group);
                System.out.println("品番 : " + itemNumber);
                System.out.println("品名 : " + itemName);
                System.out.println("当週売点 : " + numSalesPoint);
                System.out.println("----------------------");
            }

            //ブロック順位3位以上かつ全社順位10位以下かつ自店順位3位以上（地域特性）
            if (numRankBlock <= 3.0 && numRankCompany >= 6.0 && numRankStore <= 3.0) {
                //取得した文字列の表示
                System.out.println("地域特性アイテム");
                System.out.println("全社順位 : " + numRankCompany);
                System.out.println("ブロック順位 : " + numRankBlock);
                System.out.println("自店順位 : " + numRankStore);
                System.out.println("部門 : " + Group);
                System.out.println("品番 : " + itemNumber);
                System.out.println("品名 : " + itemName);
                System.out.println("当週売点 : " + numSalesPoint);
                System.out.println("----------------------");

            }
            //~~~~ここまでを繰り返し（当週売れ点ゼロはcontinu処理）~~~~//
        }
    }

/*
    static public String bestSeller() throws EncryptedDocumentException, IOException{
        //エクセルファイルへアクセス
        Workbook excel_bs = WorkbookFactory.create(new File("C://tenpo_hinban_jisseki_1319_202230.xlsx"));
        // エクセルシート名
        Sheet sheet = excel_bs.getSheet("店別");

        /*~~~~ここから~~~~*/
/*        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
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


            /*
            record ItemInfo(double numRankCompany, double numRankBlock, double numRankStore, String Group,
                                        String itemNumber, String itemName, double numSalesPoint){}
            new itemInfo()
            */


            //全社順位3位以上かつ自店順位3位以上かつ当週売点5点以上（売れ筋）
/*            if (numSalesPoint >= 5.0 && numRankCompany <= 3 && numRankStore <= 3.0) {
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
        return "list";
    }
    static public String chanceUp() throws EncryptedDocumentException, IOException{
        //エクセルファイルへアクセス
        Workbook excel_cu = WorkbookFactory.create(new File("C://tenpo_hinban_jisseki_1319_202230.xlsx"));
        // エクセルシート名
        Sheet sheet = excel_cu.getSheet("店別");

        /*~~~~ここから~~~~*/
/*        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
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


        /*
        record ItemInfo(double numRankCompany, double numRankBlock, double numRankStore, String Group,
                                    String itemNumber, String itemName, double numSalesPoint){}
        new itemInfo()
        */


            //全社順位3位以上かつブロック順位3位以上かつ自店順位10位以下（売れ筋候補）
/*            if (numRankCompany <= 3.0 && numRankBlock <= 3.0 && numRankStore >= 10.0) {
                //取得した文字列の表示
                System.out.println("売れ筋候補アイテム");
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
        return "home";
    }
    static public String darkHorse() throws EncryptedDocumentException, IOException{
        //エクセルファイルへアクセス
        Workbook excel_dh = WorkbookFactory.create(new File("C://tenpo_hinban_jisseki_1319_202230.xlsx"));
        // エクセルシート名
        Sheet sheet = excel_dh.getSheet("店別");

        /*~~~~ここから~~~~*/
/*        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
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


        /*
        record ItemInfo(double numRankCompany, double numRankBlock, double numRankStore, String Group,
                                    String itemNumber, String itemName, double numSalesPoint){}
        new itemInfo()
        */


            //ブロック順位3位以上かつ全社順位10位以下かつ自店順位3位以上（地域特性）
/*            if (numRankBlock <= 3.0 && numRankCompany >= 6.0 && numRankStore <= 3.0) {
                //取得した文字列の表示
                System.out.println("地域特性アイテム");
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
        return "add";
    }    */





}
