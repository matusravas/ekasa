package com.mrdev;

import java.io.FileNotFoundException;
import java.io.IOException;

public class Main {

    public static void main(String[] args) {
        HyperBoostLogic hyperBoostLogic = new HyperBoostLogic();
        try {
            hyperBoostLogic.openExcelDoc();
        } catch (IOException e) {
            System.out.println("Dokument report.xlsx alebo cintorin.xlsx neexistuje v adresari, kde sa nachadza spustany .exe subor,");
            System.out.println("Excel subory sa MUSIA! volat report.xlsx a cintorin.xlsx!");
            return;
        }
        hyperBoostLogic.readDocumentItems();
        System.out.println("Uzavierka k datumu: " + hyperBoostLogic.getDateFromDocument());
        hyperBoostLogic.readCementaryItems();
        hyperBoostLogic.readInvalidItems();
        hyperBoostLogic.removeInvalidItems();
        hyperBoostLogic.getRentsDriveInsServicesGoods();

        hyperBoostLogic.sumItUp();

        try {
            hyperBoostLogic.writeDataToExcel();
        } catch (FileNotFoundException e) {
            System.out.println("Do suboru sa neda zapisovat, pretoze dokument report.xlsx je prave otovreny, zatvor ho a spusti .exe znova!");
            return;
        } catch (IOException e) {
            System.out.println("Hodnoty sa nedaju zapisat, pretoze dokument body.xlsx je prave otovreny, zatvor ho a spusti .exe znova!");
        }

        try {
            hyperBoostLogic.closeExcelDoc();
        } catch (IOException e) {
            System.out.println("Dokument report.xlsx je prave otovreny, zatvor ho a spusti ma znova !!!");
            return;
        }
        System.out.println("\nUspesne vytvoreny harok Sumar v subore report...");
        System.out.print("\nStlac Enter pre zatvorenie okna...");
        try {
            System.in.read();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
