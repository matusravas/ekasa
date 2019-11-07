package com.mrdev;

public class DocumentItem {
    private String uid;
    private String itemName;
    private Double count;
    private Double sadzbaDPH;
    private Double price;



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
