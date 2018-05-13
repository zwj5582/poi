/*
 *  Created by ZhongWenjie on 2018-05-07 14:21
 */

package com.excel.poi.properties;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

@Getter@Setter
public class CellStyleProperties {

    private BorderStyle borderTop = BorderStyle.THIN;

    private BorderStyle borderBottom = BorderStyle.THIN;

    private BorderStyle borderLeft = BorderStyle.THIN;

    private BorderStyle borderRight = BorderStyle.THIN;

    private HorizontalAlignment alignment = HorizontalAlignment.CENTER;

    private VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;

    private HSSFColor.HSSFColorPredefined topBorderColor = HSSFColor.HSSFColorPredefined.BLACK;

    private HSSFColor.HSSFColorPredefined bottomBorderColor = HSSFColor.HSSFColorPredefined.BLACK;

    private HSSFColor.HSSFColorPredefined leftBorderColor = HSSFColor.HSSFColorPredefined.BLACK;

    private HSSFColor.HSSFColorPredefined rightBorderColor = HSSFColor.HSSFColorPredefined.BLACK;

    private HSSFColor.HSSFColorPredefined fillBackgroundColor = HSSFColor.HSSFColorPredefined.WHITE;

    private String dataFormat = "";

    private boolean forceCover = false;

    private CellStyleProperties(){

    }

    public static class Builder {

        private CellStyleProperties INSTANCE = new CellStyleProperties();

        private CellStyleProperties DEFAULT = new CellStyleProperties();

        public Builder borderTop(BorderStyle borderTop) {
            INSTANCE.borderTop = borderTop;
            return this;
        }

        public Builder borderBottom(BorderStyle borderBottom) {
            INSTANCE.borderBottom = borderBottom;
            return this;
        }

        public Builder borderLeft(BorderStyle borderLeft) {
            INSTANCE.borderLeft = borderLeft;
            return this;
        }

        public Builder borderRight(BorderStyle borderRight) {
            INSTANCE.borderRight = borderRight;
            return this;
        }

        public Builder alignment(HorizontalAlignment alignment) {
            INSTANCE.alignment = alignment;
            return this;
        }

        public Builder verticalAlignment(VerticalAlignment verticalAlignment) {
            INSTANCE.verticalAlignment = verticalAlignment;
            return this;
        }

        public Builder topBorderColor(HSSFColor.HSSFColorPredefined topBorderColor) {
            INSTANCE.topBorderColor = topBorderColor;
            return this;
        }

        public Builder bottomBorderColor(HSSFColor.HSSFColorPredefined bottomBorderColor) {
            INSTANCE.bottomBorderColor = bottomBorderColor;
            return this;
        }

        public Builder leftBorderColor(HSSFColor.HSSFColorPredefined leftBorderColor) {
            INSTANCE.leftBorderColor = leftBorderColor;
            return this;
        }

        public Builder rightBorderColor(HSSFColor.HSSFColorPredefined rightBorderColor) {
            INSTANCE.rightBorderColor = rightBorderColor;
            return this;
        }

        public Builder fillBackgroundColor(HSSFColor.HSSFColorPredefined fillBackgroundColor) {
            INSTANCE.fillBackgroundColor = fillBackgroundColor;
            return this;
        }

        public Builder dataFormat(String dataFormat) {
            INSTANCE.dataFormat = dataFormat;
            return this;
        }

        public Builder forceCover(boolean forceCover) {
            INSTANCE.forceCover = forceCover;
            return this;
        }

        public CellStyleProperties build(){
            return INSTANCE;
        }

        public CellStyleProperties defaultCellStyleProperties(){
            return DEFAULT;
        }

    }

}
