package com.mrdev;

import java.io.FileNotFoundException;
import java.io.IOException;

public class Main {

    public static void main(String[] args) {
        HyperBoostLogic hyperBoostLogic = new HyperBoostLogic();
        try {
            hyperBoostLogic.openExcelDoc();
        } catch (IOException e) {
            System.out.println("Dokument report.xlsx alebo cintorin.xlsx neexistuje v adresari, kde sa nachadza spustany .exe subor," +
                    "\n alebo nie su pomenovane report.xlsx a cintorin.xlsx!");
            System.out.println("Skontroluj spravnost harkov!\n-Prvy harok doklady\n-Druhy harok polozky dokladu");
            return;
        }
        hyperBoostLogic.readDocumentItems();
        hyperBoostLogic.getDate();
        hyperBoostLogic.readCementaryItems();
        hyperBoostLogic.readInvalidItems();
        hyperBoostLogic.removeInvalidItems();
        hyperBoostLogic.getRentsDriveInsServicesGoods();

        hyperBoostLogic.sumRent();
        hyperBoostLogic.sumDriveIns();
        hyperBoostLogic.sumServices();
        hyperBoostLogic.sumGoods();

        try {
            hyperBoostLogic.writeDataToExcel();
        } catch (FileNotFoundException e) {
            System.out.println("Do suboru sa neda zapisovat, pretoze dokument report.xlsx je prave otovreny, zatvor ho a spusti .exe znova!");
//            e.printStackTrace();
            return;
        } catch (IOException e) {
            System.out.println("Hodnoty sa nedaju zapisat, pretoze dokument body.xlsx je prave otovreny, zatvor ho a spusti .exe znova!");
//            e.printStackTrace();

        }

        try {
            hyperBoostLogic.closeExcelDoc();
        } catch (IOException e) {
            System.out.println("Dokument report.xlsx je prave otovreny, zatvor ho a spusti ma znova !!!");
//            e.printStackTrace();
            return;
        }

        System.out.println("Uspesne vytvoreny harok Sumar v subore report.");
    }
}
