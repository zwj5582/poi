/*
 *  Created by ZhongWenjie on 2018-05-12 12:07
 */

package com.excel.poi.properties;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public interface PropertiesConstants {

    BorderStyle borderTop = BorderStyle.THIN;

    BorderStyle borderBottom = BorderStyle.THIN;

    BorderStyle borderLeft = BorderStyle.THIN;

    BorderStyle borderRight = BorderStyle.THIN;

    HorizontalAlignment alignment = HorizontalAlignment.CENTER;

    VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;

    HSSFColor.HSSFColorPredefined topBorderColor = HSSFColor.HSSFColorPredefined.BLACK;

    HSSFColor.HSSFColorPredefined bottomBorderColor = HSSFColor.HSSFColorPredefined.BLACK;

    HSSFColor.HSSFColorPredefined leftBorderColor = HSSFColor.HSSFColorPredefined.BLACK;

    HSSFColor.HSSFColorPredefined rightBorderColor = HSSFColor.HSSFColorPredefined.BLACK;

    HSSFColor.HSSFColorPredefined fillBackgroundColor = HSSFColor.HSSFColorPredefined.WHITE;

    String dataFormat = "";



    String FontName = "宋体";

    boolean Bold = false;

    HSSFColor.HSSFColorPredefined Color = HSSFColor.HSSFColorPredefined.BLACK;

    short FontHeightInPoints = 12;

    String EMPTY = "";

}
