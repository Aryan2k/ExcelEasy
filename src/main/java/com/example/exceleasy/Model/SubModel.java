package com.example.exceleasy.Model;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.springframework.data.annotation.Id;
import org.springframework.data.mongodb.core.mapping.Document;

import java.util.HashMap;

@Getter
@Setter
@ToString

@Document(collection = "ExcelSubInfo")
public class SubModel {
    public boolean getStatus() {
        return this.status;
    }

    public String getSOR() {
        return SOR;
    }

    public void setSOR(String SOR) {
        this.SOR = SOR;
    }

    public Double getQuantity() {
        return Quantity;
    }

    public void setQuantity(Double quantity) {
        Quantity = quantity;
    }

    public Double getAmount() {
        return Amount;
    }

    public void setAmount(Double amount) {
        Amount = amount;
    }

    @SuppressWarnings("unused")
    public boolean isStatus() {
        return status;
    }

    public void setStatus(boolean status) {
        this.status = status;
    }

    public HashMap<String, Double> getConstituentSheets() {
        return constituentSheets;
    }

    public void setConstituentSheets(HashMap<String, Double> constituentSheets) {
        this.constituentSheets = constituentSheets;
    }

    public String getDESCRIPTION_OF_ITEMS() {
        return DESCRIPTION_OF_ITEMS;
    }

    public void setDESCRIPTION_OF_ITEMS(String DESCRIPTION_OF_ITEMS) {
        this.DESCRIPTION_OF_ITEMS = DESCRIPTION_OF_ITEMS;
    }

    @Id
    String SOR;
    Double Quantity;
    Double Amount;
    boolean status; //determines if the item in present in the main repo or not.
    HashMap<String, Double> constituentSheets;
    String DESCRIPTION_OF_ITEMS;   //only for those items which don't exist in the main db.

}

