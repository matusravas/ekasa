package com.mrdev;

import java.util.ArrayList;

public class Data {
    private ArrayList<DocumentItem> documentItems = new ArrayList<>();
    private ArrayList<String> invalidItems = new ArrayList<>();


    private static Data INSTANCE;

    public static Data getInstance() {
        if (INSTANCE == null) {
            INSTANCE = new Data();
            return INSTANCE;
        } else return INSTANCE;
    }

    public ArrayList<DocumentItem> getDocumentItems() {
        return documentItems;
    }

    public ArrayList<String> getInvalidItems() {
        return invalidItems;
    }

    public void setDocumentItems(ArrayList<DocumentItem> documentItems) {
        this.documentItems = documentItems;
    }

    public void setInvalidItems(ArrayList<String> invalidItems) {
        this.invalidItems = invalidItems;
    }

}
