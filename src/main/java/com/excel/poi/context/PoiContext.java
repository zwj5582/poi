/*
 *  Created by ZhongWenjie on 2018-05-07 16:24
 */

package com.excel.poi.context;

import com.excel.poi.support.CONTENT_TYPE;
import com.excel.poi.chooser.CellStyleChooser;
import com.excel.poi.render.CellStyleReRender;
import com.excel.poi.properties.*;
import com.excel.poi.annotation.ConditionalStyle;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Getter@Setter
@NoArgsConstructor
public class PoiContext {

    private final static String FONT_PROPERTIES_CLASS = FontProperties.class.getName();

    private final static String CELLSTYLE_PROPERTIES_CLASS = CellStyleProperties.class.getName();

    private Map<String,String> titleMapping = new LinkedHashMap<String,String>();

    private List<String> expressions = new ArrayList<String>();

    private GlobalProperties globalProperties;

    private CellStyleProperties globalPerCellStyleProperties;

    private CellStyleProperties globalHeaderPerCellStyleProperties;

    private CellStyleProperties globalBodyPerCellStyleProperties;

    private FontProperties globalFontProperties;

    private FontProperties bodyFontProperties;

    private FontProperties headFontProperties;

    private EntityProperties entityProperties;

    private Map<String,PerCellStyleHolder> perCellStyleHolderMap = new LinkedHashMap<String,PerCellStyleHolder>();

    private HSSFWorkbook workbook;

    private String fetchFieldNameByIndex(int index) {
        String titleName = globalProperties.getTitleNamesByIndex(index);
        String fieldName = titleMapping.get(titleName);
        if ( fieldName == null ) return PropertiesConstants.EMPTY;
        return fieldName;
    }

    public CellType fetchCellType(int index) {
        String fieldName = fetchFieldNameByIndex(index);
        return entityProperties.get(fieldName);
    }

    public ConditionalStyle fetchConditionalStyle(Object val, HSSFCellStyle cellStyle, PerCellStyleHolder perCellStyleHolder, int col_index) {
        CellStyleReRender render = perCellStyleHolder.getRender();
        render.reRender(workbook, cellStyle, val, col_index);
        CellStyleChooser chooser = perCellStyleHolder.getChooser();
        return chooser.fetchCellStyle(val, col_index);
    }


    public HSSFCellStyle createCellStyle(CONTENT_TYPE type, int index) {

        if (CONTENT_TYPE.TITLE.equals(type)) {
            HSSFCellStyle cellStyle = workbook.createCellStyle();
            CellStyleProperties cellStyleProperties = null;
            FontProperties fontProperties = null;

            cellStyleProperties = doChoose(globalProperties.getTitleCellStyleProperties(),
                    null,
                    null,
                    CellStyleProperties.class
            );
            fontProperties = doChoose(globalProperties.getTitleFontProperties(),
                    null,
                    null,
                    FontProperties.class
            );

            assert cellStyleProperties != null;
            fillCellStyle(cellStyle,cellStyleProperties);

            assert fontProperties != null;
            fillFont(cellStyle,fontProperties);

            return cellStyle;

        }

        PerCellStyleHolder perCellStyleHolder = fetchPerCellStyleHolder(index);

        HSSFCellStyle cellStyle = workbook.createCellStyle();

        CellStyleProperties cellStyleProperties = null;
        FontProperties fontProperties = null;

        if ( CONTENT_TYPE.HEADER.equals(type) ) {
            cellStyleProperties = doChoose(perCellStyleHolder.getPerHeaderCellStyleProperties(),
                    globalHeaderPerCellStyleProperties,
                    globalPerCellStyleProperties,
                    CellStyleProperties.class
            );
            fontProperties = doChoose(perCellStyleHolder.getHeadFontProperties(),
                    headFontProperties,
                    globalFontProperties,
                    FontProperties.class
            );
        }else if ( CONTENT_TYPE.BODY.equals(type) ){
            cellStyleProperties = doChoose(perCellStyleHolder.getPerBodyCellStyleProperties(),
                    globalBodyPerCellStyleProperties,
                    globalPerCellStyleProperties,
                    CellStyleProperties.class
            );
            fontProperties = doChoose(perCellStyleHolder.getBodyFontProperties(),
                    bodyFontProperties,
                    globalFontProperties,
                    FontProperties.class
            );
        }

        assert cellStyleProperties != null;
        fillCellStyle(cellStyle,cellStyleProperties);

        assert fontProperties != null;
        fillFont(cellStyle,fontProperties);

        return cellStyle;
    }

    private <T> T doChoose(Object per, Object type, Object global, Class<T> clz ) {
        return ( CELLSTYLE_PROPERTIES_CLASS.equals(clz.getName()) ) ?
                clz.cast(doChoose((CellStyleProperties) per, (CellStyleProperties) type, (CellStyleProperties) global)) :
                ( FONT_PROPERTIES_CLASS.equals(clz.getName()) ) ?
                        clz.cast(doChoose((FontProperties) per, (FontProperties) type, (FontProperties) global)) :
                        null;
    }

    private CellStyleProperties doChoose(CellStyleProperties per, CellStyleProperties type, CellStyleProperties global) {
        CellStyleProperties defaultCellStyleProperties = new CellStyleProperties.Builder().defaultCellStyleProperties();
        defaultCellStyleProperties = ( global == null ) ? defaultCellStyleProperties :
                ( global.isForceCover() ) ? global :
                        coverProperties(defaultCellStyleProperties,global);
        defaultCellStyleProperties = ( type == null ) ? defaultCellStyleProperties :
                ( type.isForceCover() ) ? global :
                        coverProperties(defaultCellStyleProperties,type);
        defaultCellStyleProperties = ( per == null ) ? defaultCellStyleProperties :
                ( per.isForceCover() ) ? global :
                        coverProperties(defaultCellStyleProperties,per);
        return defaultCellStyleProperties;
    }

    private FontProperties doChoose(FontProperties per, FontProperties type, FontProperties global) {
        FontProperties defaultCellStyleProperties = new FontProperties();
        defaultCellStyleProperties = ( global == null ) ? defaultCellStyleProperties :
                ( global.isForceCover() ) ? global :
                        coverProperties(defaultCellStyleProperties,global);
        defaultCellStyleProperties = ( type == null ) ? defaultCellStyleProperties :
                ( type.isForceCover() ) ? global :
                        coverProperties(defaultCellStyleProperties,type);
        defaultCellStyleProperties = ( per == null ) ? defaultCellStyleProperties :
                ( per.isForceCover() ) ? global :
                        coverProperties(defaultCellStyleProperties,per);
        return defaultCellStyleProperties;
    }

    private CellStyleProperties coverProperties(CellStyleProperties defaultCellStyleProperties, CellStyleProperties properties) {
        if (!properties.getBorderTop().equals(PropertiesConstants.borderTop))
            defaultCellStyleProperties.setBorderTop(properties.getBorderTop());
        if (!properties.getBorderBottom().equals(PropertiesConstants.borderBottom))
            defaultCellStyleProperties.setBorderBottom(properties.getBorderBottom());
        if (!properties.getBorderLeft().equals(PropertiesConstants.borderLeft))
            defaultCellStyleProperties.setBorderLeft(properties.getBorderLeft());
        if (!properties.getBorderRight().equals(PropertiesConstants.borderRight))
            defaultCellStyleProperties.setBorderRight(properties.getBorderRight());
        if (!properties.getAlignment().equals(PropertiesConstants.alignment))
            defaultCellStyleProperties.setAlignment(properties.getAlignment());
        if (!properties.getVerticalAlignment().equals(PropertiesConstants.verticalAlignment))
            defaultCellStyleProperties.setVerticalAlignment(properties.getVerticalAlignment());
        if (!properties.getTopBorderColor().equals(PropertiesConstants.topBorderColor))
            defaultCellStyleProperties.setTopBorderColor(properties.getTopBorderColor());
        if (!properties.getBottomBorderColor().equals(PropertiesConstants.bottomBorderColor))
            defaultCellStyleProperties.setBottomBorderColor(properties.getBottomBorderColor());
        if (!properties.getLeftBorderColor().equals(PropertiesConstants.leftBorderColor))
            defaultCellStyleProperties.setLeftBorderColor(properties.getLeftBorderColor());
        if (!properties.getRightBorderColor().equals(PropertiesConstants.rightBorderColor))
            defaultCellStyleProperties.setRightBorderColor(properties.getRightBorderColor());
        if (!properties.getFillBackgroundColor().equals(PropertiesConstants.fillBackgroundColor))
            defaultCellStyleProperties.setFillBackgroundColor(properties.getFillBackgroundColor());
        if (!properties.getDataFormat().equals(PropertiesConstants.dataFormat))
            defaultCellStyleProperties.setDataFormat(properties.getDataFormat());

        return defaultCellStyleProperties;
    }

    private FontProperties coverProperties(FontProperties defaultFontProperties, FontProperties properties) {
        if (!properties.getFontName().equals(PropertiesConstants.FontName))
            defaultFontProperties.setFontName(properties.getFontName());
        if (properties.isBold() != PropertiesConstants.Bold)
            defaultFontProperties.setBold(properties.isBold());
        if (!properties.getColor().equals(PropertiesConstants.Color))
            defaultFontProperties.setColor(properties.getColor());
        if (properties.getFontHeightInPoints() != PropertiesConstants.FontHeightInPoints)
            defaultFontProperties.setFontHeightInPoints(properties.getFontHeightInPoints());

        return defaultFontProperties;
    }

    private void fillCellStyle(HSSFCellStyle cellStyle,CellStyleProperties perHeaderCellStyleProperties) {
        if (!perHeaderCellStyleProperties.getDataFormat().equals(""))
            cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(perHeaderCellStyleProperties.getDataFormat()));
        cellStyle.setFillForegroundColor(perHeaderCellStyleProperties.getFillBackgroundColor().getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setVerticalAlignment(perHeaderCellStyleProperties.getVerticalAlignment());
        cellStyle.setAlignment(perHeaderCellStyleProperties.getAlignment());
        cellStyle.setRightBorderColor(perHeaderCellStyleProperties.getRightBorderColor().getIndex());
        cellStyle.setLeftBorderColor(perHeaderCellStyleProperties.getLeftBorderColor().getIndex());
        cellStyle.setBottomBorderColor(perHeaderCellStyleProperties.getBottomBorderColor().getIndex());
        cellStyle.setTopBorderColor(perHeaderCellStyleProperties.getTopBorderColor().getIndex());
        cellStyle.setBorderRight(perHeaderCellStyleProperties.getBorderRight());
        cellStyle.setBorderLeft(perHeaderCellStyleProperties.getBorderLeft());
        cellStyle.setBorderBottom(perHeaderCellStyleProperties.getBorderBottom());
        cellStyle.setBorderTop(perHeaderCellStyleProperties.getBorderTop());
    }

    private void fillFont(HSSFCellStyle cellStyle, FontProperties fontProperties) {
        HSSFFont font = workbook.createFont();
        font.setFontName(fontProperties.getFontName());
        font.setBold(fontProperties.isBold());
        font.setColor(fontProperties.getColor().getIndex());
        font.setFontHeightInPoints(fontProperties.getFontHeightInPoints());
        cellStyle.setFont(font);
    }

    public PerCellStyleHolder fetchPerCellStyleHolder(int index) {
        String fieldName = fetchFieldNameByIndex(index);
        return perCellStyleHolderMap.get(fieldName);
    }

}
