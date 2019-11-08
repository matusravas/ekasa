package com.mrdev;

public class DocumentItem {
    private String uid;
    private String itemName;
    private Double count;
    private Double sadzbaDPH;
    private Double price;
    private int itemType;  //0 - rent, 1 - driveIn, 2 - good, 3 - service

    public void setItemType(int itemType) {
        this.itemType = itemType;
    }

    public int getItemType() {
        return itemType;
    }

    public String getUid() {
        return uid;
    }

    public String getItemName() {
        return itemName;
    }

    public Double getCount() {
        return count;
    }

    public Double getSadzbaDPH() {
        return sadzbaDPH;
    }

    public Double getPrice() {
        return price;
    }

    public void setUid(String uid) {
        this.uid = uid;
    }

    public void setItemName(String itemName) {
        this.itemName = itemName;
    }

    public void setCount(Double count) {
        this.count = count;
    }

    public void setSadzbaDPH(Double sadzbaDPH) {
        this.sadzbaDPH = sadzbaDPH;
    }

    public void setPrice(Double price) {
        this.price = price;
    }
}
