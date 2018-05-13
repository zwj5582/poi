/*
 *  Created by ZhongWenjie on 2018-05-07 10:39
 */

package com.excel;

import com.excel.poi.support.CONTENT_TYPE;
import com.excel.poi.annotation.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;

import java.io.Serializable;
import java.util.Date;

@Entity(title = "人员信息")
@CellStyle(Type = CONTENT_TYPE.HEADER,
        FillBackgroundColor = HSSFColor.HSSFColorPredefined.GREY_50_PERCENT,
        BorderTop = BorderStyle.DOUBLE,
        BorderBottom = BorderStyle.DOUBLE,
        BorderLeft = BorderStyle.DOUBLE,
        BorderRight = BorderStyle.DOUBLE
)
@HeaderFont
public class WorkBookEntity implements Serializable {

    @DisplayName("姓名")
    @Type
    @Expression("patient.name")
    /*@CellStyles(
            value = {
                    @CellStyle(Type = CONTENT_TYPE.BODY, FillBackgroundColor=HSSFColor.HSSFColorPredefined.RED)
            }
    )*/
    private String name;

    @DisplayName("年龄")
    @Type(CellType.NUMERIC)
    @Expression("patient.age")
    /*@Conditions(values = {
            @Condition(value = "com.excel.CustomConditionalChooser",
                    True = @ConditionalStyle(Color = "RED"),
                    False = @ConditionalStyle())
    })*/
    /*@RowCondition(value = "com.excel.CustomConditionalChooser",
            True = @ConditionalStyle(FillBackgroundColor = "BLUE"),
            False = @ConditionalStyle())*/
    private Integer age;

    @DisplayName("性别")
    @Expression("patient.sex")
    private String sex;

    @DisplayName("地址")
    @Type
    @Expression("address")
    @CellStyle(Type = CONTENT_TYPE.ALL,ForceCover = true)
    private String address;

    @DisplayName("日期")
    @Type(CellType.NUMERIC)
    @CellStyle(Type = CONTENT_TYPE.BODY,DataFormat = "m/d/yy h:mm")
    @Expression("patient.date")
    private Date date;

    public WorkBookEntity() {
    }

    public WorkBookEntity(String name, Integer age, String sex, String address, Date date) {
        this.name = name;
        this.age = age;
        this.sex = sex;
        this.address = address;
        this.date = date;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }
}
