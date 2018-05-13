/*
 *  Created by ZhongWenjie on 2018-05-11 10:50
 */

package com.excel.poi.render;

import com.excel.poi.annotation.Condition;
import com.excel.poi.annotation.ConditionalStyle;
import com.excel.poi.chooser.CellStyleChooser;
import com.excel.poi.chooser.ConditionalChooser;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class CellStyleReRender {

    private List<CellStyleChooser> choosers = new ArrayList<CellStyleChooser>();

    private List<ConditionalStyle> styleList = new ArrayList<ConditionalStyle>();

    public CellStyleReRender( Condition[] conditions ) throws IllegalAccessException, InstantiationException, ClassNotFoundException {
        if (conditions != null){
            for ( Condition condition : conditions ) {
                if ( condition != null ) {
                    String chooserClassStr = condition.value();
                    choosers.add(
                            new CellStyleChooser
                                    ((ConditionalChooser) Class.forName(chooserClassStr).newInstance(), condition.True(), condition.False())
                    );
                }
            }
        }
    }

    public CellStyleReRender( List<ConditionalStyle> styleList ) {
        this.styleList = styleList;
    }

    public HSSFCellStyle reRender(Workbook workbook, HSSFCellStyle cellStyle, Object input, int rowNum){
        if (!styleList.isEmpty()){
            for (ConditionalStyle conditionalStyle : styleList) {
                doRender(workbook, cellStyle, conditionalStyle);
            }
            return cellStyle;
        }
        for ( CellStyleChooser styleChooser : choosers ) {
            ConditionalStyle conditionalStyle = styleChooser.fetchCellStyle(input,rowNum);
            doRender(workbook, cellStyle, conditionalStyle);
        }
        return cellStyle;
    }

    private void doRender(Workbook workbook, HSSFCellStyle cellStyle, ConditionalStyle conditionalStyle) {
        if (!"".equals(conditionalStyle.BorderTop())) {
            cellStyle.setBorderTop(BorderStyle.valueOf(conditionalStyle.BorderTop()));
        }
        if (!"".equals(conditionalStyle.BorderBottom())) {
            cellStyle.setBorderBottom(BorderStyle.valueOf(conditionalStyle.BorderBottom()));
        }
        if (!"".equals(conditionalStyle.BorderLeft())){
            cellStyle.setBorderLeft(BorderStyle.valueOf(conditionalStyle.BorderLeft()));
        }
        if (!"".equals(conditionalStyle.BorderRight())) {
            cellStyle.setBorderRight(BorderStyle.valueOf(conditionalStyle.BorderRight()));
        }
        if (!"".equals(conditionalStyle.Alignment())) {
            cellStyle.setAlignment(HorizontalAlignment.valueOf(conditionalStyle.Alignment()));
        }
        if (!"".equals(conditionalStyle.VerticalAlignment())) {
            cellStyle.setVerticalAlignment(VerticalAlignment.valueOf(conditionalStyle.VerticalAlignment()));
        }
        if (!"".equals(conditionalStyle.TopBorderColor())) {
            cellStyle.setTopBorderColor(HSSFColor.HSSFColorPredefined.valueOf(conditionalStyle.TopBorderColor()).getIndex());
        }
        if (!"".equals(conditionalStyle.BottomBorderColor())) {
            cellStyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.valueOf(conditionalStyle.BottomBorderColor()).getIndex());
        }
        if (!"".equals(conditionalStyle.LeftBorderColor())) {
            cellStyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.valueOf(conditionalStyle.LeftBorderColor()).getIndex());
        }
        if (!"".equals(conditionalStyle.RightBorderColor())) {
            cellStyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.valueOf(conditionalStyle.RightBorderColor()).getIndex());
        }
        if (!"".equals(conditionalStyle.FillBackgroundColor())) {
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.valueOf(conditionalStyle.FillBackgroundColor()).getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        if (!"".equals(conditionalStyle.DataFormat())) {
            cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(conditionalStyle.DataFormat()));
        }
        HSSFFont font = cellStyle.getFont(workbook);
        if (!"".equals(conditionalStyle.FontName())) {
            font.setFontName(conditionalStyle.FontName());
        }
        if (!"".equals(conditionalStyle.Bold())) {
            font.setBold(Boolean.valueOf(conditionalStyle.Bold()));
        }
        if (!"".equals(conditionalStyle.Color())) {
            font.setColor(HSSFColor.HSSFColorPredefined.valueOf(conditionalStyle.Color()).getIndex());
        }
        if (!"".equals(conditionalStyle.FontHeightInPoints())) {
            font.setFontHeightInPoints(Short.valueOf(conditionalStyle.FontHeightInPoints()));
        }
    }



}
