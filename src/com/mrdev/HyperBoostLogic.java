package com.mrdev;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
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
    private XSSFSheet doklady;
    private XSSFSheet polozkyDokladu;
    private XSSFSheet cementary;
    private XSSFSheet exportSheet;
    private XSSFSheet infoSheet;
    private double sumServices = 0;
    private double sumDriveIns = 0;
    private double sumRent = 0;
    private double sumGoods = 0;
    private ArrayList<String> services = new ArrayList<>(); //stores all services NOT REPEATING
    private ArrayList<String> goods = new ArrayList<>();

    //FINAL ARRAYS TO BE WRITTEN IN EXCEL
    private ArrayList<DocumentItem> driveInsF = new ArrayList<>();
    private ArrayList<DocumentItem> rentsF = new ArrayList<>();
    private ArrayList<DocumentItem> servicesF = new ArrayList<>();
    private ArrayList<DocumentItem> goodsF = new ArrayList<>();


    void openExcelDoc() throws IOException {
        fileInEkasa = new FileInputStream(new File("C:\\BOX\\eKasa\\report.xlsx"));
        workbookEkasa = new XSSFWorkbook(fileInEkasa);
        fileInCementary = new FileInputStream(new File("C:\\BOX\\eKasa\\cintorin.xlsx"));
        workbookCementary = new XSSFWorkbook(fileInCementary);

        doklady = workbookEkasa.getSheetAt(0);
        System.out.println((doklady.getSheetName()));
        polozkyDokladu = workbookEkasa.getSheetAt(1);
        infoSheet = workbookEkasa.getSheetAt(2);
        cementary = workbookCementary.getSheetAt(0);
//        System.out.println("Skontroluj spravnost harkov!\n-Prvy harok doklady\n-Druhy harok polozky dokladu");
    }

    String getDate() {
        Row row = infoSheet.getRow(4);
        Date date = row.getCell(1).getDateCellValue();
        SimpleDateFormat df = new SimpleDateFormat("dd.MM.YYYY HH:mm:ss.SS");
        String fDate = df.format(date);
        System.out.println("Datum uzavierky: " + fDate);
        return fDate;
    }

    void readDocumentItems() {
        DocumentItem item;
        ArrayList<DocumentItem> items = new ArrayList<>();
        for (int i = 2; i < polozkyDokladu.getLastRowNum() + 1; i++) {
            Row row = polozkyDokladu.getRow(i);
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
        for (int i = 2; i < doklady.getLastRowNum(); i++) {
            Row row = doklady.getRow(i);
            if (row != null || !row.getCell(0).getStringCellValue().isEmpty()) {
                name = row.getCell(5).getStringCellValue().toLowerCase();
                if (name.equals("neplatný doklad")) {
                    uids.add(row.getCell(3).getStringCellValue());
                }
            }
        }
        Data.getInstance().setInvalidItems(uids);
        System.out.println("Pocet najdenych neplatnych poloziek: " + Data.getInstance().getInvalidItems().size());
    }

    void removeInvalidItems() {
        for (int i = 0; i < Data.getInstance().getDocumentItems().size(); i++) {
            for (int j = 0; j < Data.getInstance().getInvalidItems().size(); j++) {
                if (Data.getInstance().getDocumentItems().get(i).getUid().equals(Data.getInstance().getInvalidItems().get(j))) {
                    Data.getInstance().getDocumentItems().remove(i);
                    System.out.println("Zmazana neplatna polozka: " + Data.getInstance().getDocumentItems().get(i).getItemName());
                }
            }
        }

    }

    void getRentsDriveInsServicesGoods() {
        for (DocumentItem item : Data.getInstance().getDocumentItems()) {
            if (item.getItemName().contains("nájom")) {
                rentsF.add(item);
                continue;
            }
            if (item.getItemName().contains("vjazd")) {
                driveInsF.add(item);
                continue;
            }
            if (goods.indexOf(item.getItemName()) != -1 || item.getItemName().contains("PLU")) { //Ak je to tovar
                goodsF.add(item);
            } else {
                servicesF.add(item);

            }
        }

    }

    void sumRent() {
        for (DocumentItem item : rentsF) {
            sumRent += item.getPrice();
        }
        System.out.println("Suma najmy: " + sumRent + " rentsF size" + rentsF.size());
    }

    void sumDriveIns() {
        for (DocumentItem item : driveInsF) {
            sumDriveIns += item.getPrice();
        }
        System.out.println("Suma vjazdy: " + sumDriveIns + " driveInsF size " + driveInsF.size());
    }

    void sumServices() {
        for (DocumentItem item : servicesF) {
            sumServices += item.getPrice();
        }
        System.out.println("Suma sluzby: " + sumServices + " servicesF size " + servicesF.size());
    }

    void sumGoods() {
        for (DocumentItem item : goodsF) {
            sumGoods += item.getPrice();
        }
        System.out.println("Suma tovar: " + sumGoods + " goodsF size " + goodsF.size());
        System.out.println("Suma celkom: " + (sumGoods + sumServices + sumRent + sumDriveIns));
        System.out.println("Celkovy pocet poloziek: " + (goodsF.size() + servicesF.size() + rentsF.size() + driveInsF.size()));
    }

    /**
     * Stlpec za GPS data bude rozdiel GPS cas - sonar cas
     */
    void writeDataToExcel() throws IOException {
        String date = this.getDate(); //Vracia datum ku ktorej sa uzavierka viaze

        fileOut = new FileOutputStream(new File("report.xlsx"));
        removeExistingSheet(workbookEkasa);
        exportSheet = workbookEkasa.createSheet("Sumar");

        ////Zahlavie Datum
        Row row0 = exportSheet.createRow(0);
        Cell dateTitle = row0.createCell(0);
        Cell dateValue = row0.createCell(1);
        dateTitle.setCellValue("Uzavierka ku dnu");
        dateValue.setCellValue(date);

        ////Zahlavie Datum

        Row row1 = exportSheet.createRow(1);
        Cell names = row1.createCell(0);
        Cell sums = row1.createCell(1);
        Cell DPHless = row1.createCell(2);
        Cell DPH = row1.createCell(3);
        names.setCellValue("Názov");
        sums.setCellValue("Suma");
        DPHless.setCellValue("Bez DPH");
        DPH.setCellValue("DPH");

        //goods
        double goodsNoDPH = sumGoods / (1.2);
        double goodsDPH = (sumGoods / (1.2)) * 0.2;
        Row row2 = exportSheet.createRow(2);
        Cell goods = row2.createCell(0);
        Cell gSum = row2.createCell(1);
        Cell gDPHless = row2.createCell(2);
        Cell gDPH = row2.createCell(3);
        goods.setCellValue("Tovar");
        gSum.setCellValue(sumGoods);
        gDPHless.setCellValue(goodsNoDPH);
        gDPH.setCellValue(goodsDPH);

        //services
        double servicesNoDPH = sumServices / (1.2);
        double servicesDPH = (sumServices / (1.2)) * 0.2;
        Row row3 = exportSheet.createRow(3);
        Cell services = row3.createCell(0);
        Cell sSum = row3.createCell(1);
        Cell sDPHless = row3.createCell(2);
        Cell sDPH = row3.createCell(3);
        services.setCellValue("Služba");
        sSum.setCellValue(sumServices);
        sDPHless.setCellValue(servicesNoDPH);
        sDPH.setCellValue(servicesDPH);

        //DriveIns
        double driveNoDPH = sumDriveIns / (1.2);
        double driveDPH = (sumDriveIns / (1.2)) * 0.2;
        Row row5 = exportSheet.createRow(4);
        Cell drives = row5.createCell(0);
        Cell dSum = row5.createCell(1);
        Cell dDPHless = row5.createCell(2);
        Cell dDPH = row5.createCell(3);
        drives.setCellValue("Vjazdy");
        dSum.setCellValue(sumDriveIns);
        dDPHless.setCellValue(driveNoDPH);
        dDPH.setCellValue((sumDriveIns / (1.2)) * 0.2);

        Row row6 = exportSheet.createRow(5);
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

        //Rents do not count DPH
//        double rentsNoDPH = sumRent / (1.2);
//        double rentsDPH = (sumRent / (1.2)) * 0.2;
        Row row4 = exportSheet.createRow(8);
        Cell rents = row4.createCell(0);
        Cell rSum = row4.createCell(1);
        Cell rDPHless = row4.createCell(2);
        Cell rDPH = row4.createCell(3);
        rents.setCellValue("Nájmy");
        rSum.setCellValue(sumRent);
        rDPHless.setCellValue(sumRent);
        rDPH.setCellValue(0);

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