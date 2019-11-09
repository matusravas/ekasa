package com.mrdev;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class HyperBoostLogic {
    private FileInputStream fileInEkasa;
    private FileInputStream fileInCementary;
    private FileOutputStream fileOut;
    private XSSFWorkbook workbookEkasa;
    private XSSFWorkbook workbookCementary;
    private XSSFSheet documentsSheet;
    private XSSFSheet documentItemsSheet;
    private XSSFSheet cementary;
    private XSSFSheet exportSheet;
    private XSSFSheet infoSheet;

    private File reportFile;
    private File cementaryFile;

    private double sumServices = 0;
    private double sumDriveIns = 0;
    private double sumRent = 0;
    private double sumGoods = 0;
    private ArrayList<String> services = new ArrayList<>(); //stores all services
    private ArrayList<String> goods = new ArrayList<>(); // stores all goods

    //CATEGORIZED LIST OF ALL ITEMS, each item has its type available item.getItemType()
    // 0 - rent, 1 - driveIn, 2 - good, 3 - service, 4 - uncategorized
    private ArrayList<DocumentItem> categorizedItems = new ArrayList<>();
    private ArrayList<DocumentItem> unCategorizedItems = new ArrayList<>();

    void findExcelFiles() throws FileNotFoundException {
        File folder = new File("./");
        File[] listOfFiles = folder.listFiles();
        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].getName().contains(".xlsx")) {
                if (listOfFiles[i].isFile() && listOfFiles[i].getName().contains("report")) {
                    reportFile = new File(listOfFiles[i].getPath());
                    fileInEkasa = new FileInputStream(reportFile);
                }
                if (listOfFiles[i].isFile() && listOfFiles[i].getName().contains("cintorin")) {
                    cementaryFile = new File(listOfFiles[i].getPath());
                    fileInCementary = new FileInputStream(cementaryFile);
                }
            }
        }
    }

    void openExcelDoc() throws IOException {
        findExcelFiles();
//        fileInEkasa = new FileInputStream(new File(".\\report.xlsx"));
        workbookEkasa = new XSSFWorkbook(fileInEkasa);
//        fileInCementary = new FileInputStream(new File(".\\cintorin.xlsx"));
        workbookCementary = new XSSFWorkbook(fileInCementary);

        documentsSheet = workbookEkasa.getSheetAt(0);
        System.out.println("Citanie dokladov z eKasy...");
        documentItemsSheet = workbookEkasa.getSheetAt(1);
        infoSheet = workbookEkasa.getSheetAt(2);
        cementary = workbookCementary.getSheetAt(0);
    }

    String getDateFromDocument() {
        Row row = infoSheet.getRow(4);
        Date date = row.getCell(1).getDateCellValue();
        SimpleDateFormat df = new SimpleDateFormat("dd.MM.YYYY HH:mm:ss.SS");
        String fDate = df.format(date);
        return fDate;
    }

    void readDocumentItems() {
        DocumentItem item;
        ArrayList<DocumentItem> items = new ArrayList<>();
        for (int i = 2; i < documentItemsSheet.getLastRowNum() + 1; i++) {
            Row row = documentItemsSheet.getRow(i);
            if (row != null || !row.getCell(0).getStringCellValue().isEmpty()) {
                item = new DocumentItem();
                item.setUid(row.getCell(0).getStringCellValue());
                item.setItemName(row.getCell(1).getStringCellValue());
                item.setCount(row.getCell(2).getNumericCellValue());
                item.setSadzbaDPH(row.getCell(3).getNumericCellValue());
                item.setPrice(row.getCell(5).getNumericCellValue());
                items.add(item);
            }
        }
        Data.getInstance().setDocumentItems(items);
    }

    void readCementaryItems() {
        for (int i = 1; i < cementary.getLastRowNum() + 1; i++) {
            Row row = cementary.getRow(i);
            if (row != null || !row.getCell(0).getStringCellValue().isEmpty()) {
                if (row.getCell(6).getNumericCellValue() == 2.0) { //sluzba
                    services.add(row.getCell(2).getStringCellValue());
                }
                if (row.getCell(6).getNumericCellValue() == 1.0) { //tovar
                    goods.add(row.getCell(2).getStringCellValue());
                }
            }
        }
    }

    void readInvalidItems() {
        String name;
        ArrayList<String> uids = new ArrayList<>();
        for (int i = 2; i < documentsSheet.getLastRowNum() + 1; i++) {
            Row row = documentsSheet.getRow(i);
            if (row != null || !row.getCell(0).getStringCellValue().isEmpty()) {
                name = row.getCell(5).getStringCellValue().toLowerCase();
                if (name.equals("neplatný doklad")) {
                    uids.add(row.getCell(3).getStringCellValue());
                }
            }
        }
        Data.getInstance().setInvalidItems(uids);
        System.out.println("\nPocet najdenych neplatnych poloziek: " + Data.getInstance().getInvalidItems().size());
    }

    void removeInvalidItems() {
        for (int i = 0; i < Data.getInstance().getDocumentItems().size(); i++) {
            for (int j = 0; j < Data.getInstance().getInvalidItems().size(); j++) {
                if (Data.getInstance().getDocumentItems().get(i).getUid().equals(Data.getInstance().getInvalidItems().get(j))) {
                    System.out.println("Vynechana neplatna polozka: " + Data.getInstance().getDocumentItems().get(i).getItemName());
                    Data.getInstance().getDocumentItems().remove(i);
                }
            }
        }
        System.out.print("\n");
    }

    void getRentsDriveInsServicesGoods() {
        for (DocumentItem item : Data.getInstance().getDocumentItems()) {
            if (item.getItemName().contains("nájom")) {
                item.setItemType(0);
            } else if (item.getItemName().contains("vjazd")) {
                item.setItemType(1);
            } else if (goods.indexOf(item.getItemName()) != -1 || item.getItemName().toLowerCase().contains("plu")) { //Ak je to tovar
                item.setItemType(2);
            } else if (services.indexOf(item.getItemName()) != -1) {
                item.setItemType(3);
            } else {
                unCategorizedItems.add(item); //na zaver su vsetky vypisane do excelu
                item.setItemType(4);
            }
            /**
             * Ked sa nenachadza item ani v nacitanych servisoch ani goods z cintorin.xlsx
             * zarad ho do others a nakonci vsetky others vypises
             */
            categorizedItems.add(item);
        }

    }

    void sumItUp() {
        for (DocumentItem item : categorizedItems) {
            switch (item.getItemType()) {
                case 0:
                    sumRent += item.getPrice();
                    break;
                case 1:
                    sumDriveIns += item.getPrice();
                    break;
                case 2:
                    sumGoods += item.getPrice();
                    break;
                case 3:
                    sumServices += item.getPrice();
                    break;
                case 4:
                    System.out.println("Nekategorizovana polozka: " + item.getItemName());
                    System.out.println("Suma: " + item.getPrice());
                    break;
            }
        }
        System.out.println("\nSuma tovar: " + sumGoods);
        System.out.println("Suma sluzby: " + sumServices);
        System.out.println("Suma vjazdy: " + sumDriveIns);
        System.out.println("Suma najmy: " + sumRent);
        System.out.println("Suma celkom: " + (sumGoods + sumServices + sumRent + sumDriveIns));
    }

    void writeDataToExcel() throws IOException {
        String date = this.getDateFromDocument(); //Vracia datum ku ktorej sa uzavierka viaze

        fileOut = new FileOutputStream(new File("report.xlsx"));
        removeExistingSheet(workbookEkasa);
        exportSheet = workbookEkasa.createSheet("Sumar");

        // Header Date
        Row row0 = exportSheet.createRow(0);
        Cell dateTitle = row0.createCell(0);
        Cell dateValue = row0.createCell(1);
        dateTitle.setCellValue("Uzávierka ku dňu");
        dateValue.setCellValue(date);

        // Header Titles
        Row row1 = exportSheet.createRow(2);
        Cell names = row1.createCell(0);
        Cell sums = row1.createCell(1);
        Cell DPHless = row1.createCell(2);
        Cell DPH = row1.createCell(3);
        names.setCellValue("Názov");
        sums.setCellValue("Suma");
        DPHless.setCellValue("Bez DPH");
        DPH.setCellValue("DPH");

        // Goods
        double goodsNoDPH = sumGoods / (1.2);
        double goodsDPH = (sumGoods / (1.2)) * 0.2;
        Row row2 = exportSheet.createRow(3);
        Cell goods = row2.createCell(0);
        Cell gSum = row2.createCell(1);
        Cell gDPHless = row2.createCell(2);
        Cell gDPH = row2.createCell(3);
        goods.setCellValue("Tovar");
        gSum.setCellValue(sumGoods);
        gDPHless.setCellValue(goodsNoDPH);
        gDPH.setCellValue(goodsDPH);

        // Services
        double servicesNoDPH = sumServices / (1.2);
        double servicesDPH = (sumServices / (1.2)) * 0.2;
        Row row3 = exportSheet.createRow(4);
        Cell services = row3.createCell(0);
        Cell sSum = row3.createCell(1);
        Cell sDPHless = row3.createCell(2);
        Cell sDPH = row3.createCell(3);
        services.setCellValue("Služba");
        sSum.setCellValue(sumServices);
        sDPHless.setCellValue(servicesNoDPH);
        sDPH.setCellValue(servicesDPH);

        // DriveIns
        double driveNoDPH = sumDriveIns / (1.2);
        double driveDPH = (sumDriveIns / (1.2)) * 0.2;
        Row row5 = exportSheet.createRow(5);
        Cell drives = row5.createCell(0);
        Cell dSum = row5.createCell(1);
        Cell dDPHless = row5.createCell(2);
        Cell dDPH = row5.createCell(3);
        drives.setCellValue("Vjazdy");
        dSum.setCellValue(sumDriveIns);
        dDPHless.setCellValue(driveNoDPH);
        dDPH.setCellValue((sumDriveIns / (1.2)) * 0.2);

        // Total
        Row row6 = exportSheet.createRow(6);
        Cell total = row6.createCell(0);
        Cell tSum = row6.createCell(1);
        Cell tDPHless = row6.createCell(2);
        Cell tDPH = row6.createCell(3);
        total.setCellValue("Spolu");
        tSum.setCellValue((sumDriveIns + sumServices + sumRent + sumGoods));
        double sumNoDPH = (driveNoDPH + servicesNoDPH + goodsNoDPH);
        tDPHless.setCellValue(sumNoDPH);
        double sumDPH = (driveDPH + servicesDPH + goodsDPH);
        tDPH.setCellValue(sumDPH);

        // Rents do not count DPH
        Row row4 = exportSheet.createRow(8);
        Cell rents = row4.createCell(0);
        Cell rSum = row4.createCell(1);
        Cell rDPHless = row4.createCell(2);
        Cell rDPH = row4.createCell(3);
        rents.setCellValue("Nájmy");
        rSum.setCellValue(sumRent);
        rDPHless.setCellValue(sumRent);

        // Header uncategorized items
        Row row7 = exportSheet.createRow(10);
        Cell uncategorized = row7.createCell(0);
        uncategorized.setCellValue("Nekategorizované položky");
        Row row8 = exportSheet.createRow(11);
        Cell cItem = row8.createCell(0);
        cItem.setCellValue("Položka");
        Cell cPrice = row8.createCell(1);
        cPrice.setCellValue("Cena celkom");

        // Listing of uncategorized items
        for (int i = 0; i < unCategorizedItems.size(); i++) {
            Row row = exportSheet.createRow(12 + i);
            Cell cellItem = row.createCell(0);
            Cell cellPrice = row.createCell(1);
            cellItem.setCellValue(unCategorizedItems.get(i).getItemName());
            cellPrice.setCellValue(unCategorizedItems.get(i).getPrice());
        }

        workbookEkasa.write(fileOut);
        fileOut.close();
    }

    private void removeExistingSheet(XSSFWorkbook workbook) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet tmpSheet = workbook.getSheetAt(i);
            if (tmpSheet.getSheetName().equals("Sumar")) {
                workbook.removeSheetAt(i);
            }
        }
    }

    void closeExcelDoc() throws IOException {
        fileInEkasa.close();
        fileInCementary.close();
    }
}
