package com.example.exceleasy.Model;


import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.springframework.data.annotation.Id;
import org.springframework.data.mongodb.core.mapping.Document;

@Getter
@Setter
@ToString

@Document(collection = "Excel")
public class Model {

    @Id
    String SOR;
    String DESCRIPTION_OF_ITEMS;
    String Unit;
    Double Rate;

    public String getSOR() {
        return SOR;
    }

    public void setSOR(String SOR) {
        this.SOR = SOR;
    }

    public String getDESCRIPTION_OF_ITEMS() {
        return DESCRIPTION_OF_ITEMS;
    }

    public void setDESCRIPTION_OF_ITEMS(String DESCRIPTION_OF_ITEMS) {
        this.DESCRIPTION_OF_ITEMS = DESCRIPTION_OF_ITEMS;
    }

    public String getUnit() {
        return Unit;
    }

    public void setUnit(String unit) {
        Unit = unit;
    }

    public Double getRate() {
        return Rate;
    }

    public void setRate(Double rate) {
        Rate = rate;
    }
}
