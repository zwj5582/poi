/*
 *  Created by ZhongWenjie on 2018-05-09 21:13
 */

package com.excel.poi.reader;

import com.excel.poi.annotation.Conditions;
import com.excel.poi.context.PerCellStyleHolder;
import com.excel.poi.context.PoiContext;
import com.excel.poi.properties.*;
import com.excel.poi.annotation.*;
import com.excel.poi.support.CONTENT_TYPE;
import com.excel.poi.chooser.CellStyleChooser;
import com.excel.poi.render.CellStyleReRender;
import com.excel.poi.chooser.ConditionalChooser;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

public class DefaultPoiPropertiesReader implements PoiPropertiesReader {

    @Override
    public PoiContext readProperties(Class<?> entityClass) throws InstantiationException, IllegalAccessException, ClassNotFoundException {

        PoiContext poiContext = new PoiContext();

        poiContext.setGlobalHeaderPerCellStyleProperties(createCellStyleProperties(entityClass, CONTENT_TYPE.HEADER));

        poiContext.setGlobalBodyPerCellStyleProperties(createCellStyleProperties(entityClass, CONTENT_TYPE.BODY));

        poiContext.setGlobalPerCellStyleProperties(createCellStyleProperties(entityClass, CONTENT_TYPE.ALL));

        Entity entity = entityClass.getAnnotation(Entity.class);
        GlobalProperties globalProperties = new GlobalProperties();
        EntityProperties entityProperties = new EntityProperties();
        globalProperties.put("title",entity.title());
        globalProperties.put("columnWidth",entity.columnWidth());
        globalProperties.put("columnHeight",entity.columnHeight());
        globalProperties.put("titleColumnHeight",entity.titleColumnHeight());

        fillTitleProperties(entity,globalProperties);

        poiContext.setGlobalProperties(globalProperties);
        poiContext.setEntityProperties(entityProperties);


        Field[] fields = entityClass.getDeclaredFields();
        List<String> titleNames = globalProperties.getTitleNames();

        List<String> expressions = poiContext.getExpressions();

        poiContext.setHeadFontProperties(createFontProperties(entityClass,CONTENT_TYPE.HEADER));
        poiContext.setBodyFontProperties(createFontProperties(entityClass,CONTENT_TYPE.BODY));
        poiContext.setGlobalFontProperties(createFontProperties(entityClass,CONTENT_TYPE.ALL));

        Map<String, PerCellStyleHolder> cellStyleHolderMap = poiContext.getPerCellStyleHolderMap();

        Map<String, String> titleMapping = poiContext.getTitleMapping();

        for (Field field : fields){

            PerCellStyleHolder perCellStyleHolder = new PerCellStyleHolder();
            perCellStyleHolder.setHeadFontProperties(createFontProperties(field, CONTENT_TYPE.HEADER));
            perCellStyleHolder.setBodyFontProperties(createFontProperties(field, CONTENT_TYPE.BODY));
            perCellStyleHolder.setPerHeaderCellStyleProperties(createCellStyleProperties(field,CONTENT_TYPE.HEADER));
            perCellStyleHolder.setPerBodyCellStyleProperties(createCellStyleProperties(field,CONTENT_TYPE.BODY));


            cellStyleHolderMap.put(field.getName(),perCellStyleHolder);

            Type type = field.getAnnotation(Type.class);
            boolean flag = (type == null) ?
                    entityProperties.put(field.getName(), CellType.STRING) :
                    entityProperties.put(field.getName(), type.value());
            if (!flag) continue;
            DisplayName displayName = field.getAnnotation(DisplayName.class);
            String name = displayName.value();
            if (name.equals(""))
                name = field.getName();
            titleNames.add(name);
            titleMapping.put(name,field.getName());
            Expression expression = field.getAnnotation(Expression.class);
            if (expression == null)
                expressions.add(field.getName());
            else
                expressions.add(expression.value());

            Conditions conditions = field.getAnnotation(Conditions.class);
            Condition[] conditionArr =
                    ( conditions != null ) ? conditions.values() : new Condition[]{};
            CellStyleReRender render = new CellStyleReRender(conditionArr);
            perCellStyleHolder.setRender(render);

            RowCondition rowCondition = field.getAnnotation(RowCondition.class);
            CellStyleChooser chooser = ( rowCondition == null ) ?
                    new CellStyleChooser(CellStyleChooser.ALWAYS_TRUE, CellStyleChooser.EMPTY, CellStyleChooser.EMPTY) :
                    new CellStyleChooser((ConditionalChooser) Class.forName(rowCondition.value()).newInstance(),
                            rowCondition.True(),
                            rowCondition.False()
                            );
            perCellStyleHolder.setChooser(chooser);
        }
        return poiContext;
    }

    private void fillTitleProperties(Entity entity, GlobalProperties globalProperties) {
        CellStyle style = entity.titleStyle();
        globalProperties.setTitleCellStyleProperties(
                new CellStyleProperties.Builder()
                        .borderTop(style.BorderTop())
                        .borderBottom(style.BorderBottom())
                        .borderLeft(style.BorderLeft())
                        .borderRight(style.BorderRight())
                        .topBorderColor(style.TopBorderColor())
                        .bottomBorderColor(style.BottomBorderColor())
                        .leftBorderColor(style.LeftBorderColor())
                        .rightBorderColor(style.RightBorderColor())
                        .alignment(style.Alignment())
                        .verticalAlignment(style.VerticalAlignment())
                        .fillBackgroundColor(style.FillBackgroundColor())
                        .dataFormat(style.DataFormat())
                        .forceCover(style.ForceCover())
                        .build()
        );
        Font font = entity.titleFont();
        globalProperties.setTitleFontProperties(
                new FontProperties
                        (font.FontName(),font.Bold(),font.Color(),font.FontHeightInPoints(),font.ForceCover())
        );
    }

    private FontProperties createFontPropertiesWithClass(Class entityClass, CONTENT_TYPE type){
        FontProperties fontProperties = null;
        switch (type){
            case HEADER:
                HeaderFont headerFont = (HeaderFont) entityClass.getAnnotation(HeaderFont.class);
                fontProperties = ( headerFont == null ) ?
                        null :
                        new FontProperties
                                (headerFont.FontName(),headerFont.Bold(),headerFont.Color(),headerFont.FontHeightInPoints(),headerFont.ForceCover())
                ;
                break;
            case BODY:
                Font font = (Font) entityClass.getAnnotation(Font.class);
                fontProperties =
                        ( font == null ) ?
                                null :
                                new FontProperties
                                        (font.FontName(),font.Bold(),font.Color(),font.FontHeightInPoints(),font.ForceCover())
                ;
                break;
            case ALL:
                GlobalFont globalFont = (GlobalFont) entityClass.getAnnotation(GlobalFont.class);
                fontProperties =
                        ( globalFont == null ) ?
                                null :
                                new FontProperties
                                        (globalFont.FontName(),globalFont.Bold(),globalFont.Color(),globalFont.FontHeightInPoints(),globalFont.ForceCover())
                ;
                break;
        }
        return fontProperties;
    }

    private FontProperties createFontPropertiesWithField(Field field, CONTENT_TYPE type){
        FontProperties fontProperties = null;
        switch (type){
            case HEADER:
                HeaderFont headerFont = field.getAnnotation(HeaderFont.class);
                if ( headerFont != null )
                    fontProperties =
                            new FontProperties
                                    (headerFont.FontName(),headerFont.Bold(),headerFont.Color(),headerFont.FontHeightInPoints(),headerFont.ForceCover())
                    ;
                break;
            case BODY:
                Font font = field.getAnnotation(Font.class);
                if ( font != null )
                    fontProperties =
                            new FontProperties
                                    (font.FontName(),font.Bold(),font.Color(),font.FontHeightInPoints(),font.ForceCover())
                    ;
                break;
        }
        return fontProperties;
    }

    private FontProperties createFontProperties(Object obj, CONTENT_TYPE type) {
        if (obj instanceof Class)
            return createFontPropertiesWithClass((Class) obj, type);
        else if (obj instanceof Field)
            return createFontPropertiesWithField((Field) obj, type);
        return null;
    }

    private CellStyleProperties createCellStyleProperties(Object obj,CONTENT_TYPE type) {

        CellStyle style = ( obj instanceof Class ) ?
                (CellStyle) ((Class)obj).getAnnotation(CellStyle.class) :
                ( obj instanceof Field ) ?
                        ((Field)obj).getAnnotation(CellStyle.class) :
                        null
                ;
        if (style != null && style.Type().equals(CONTENT_TYPE.ALL) && CONTENT_TYPE.ALL.equals(type) ){
            return new CellStyleProperties.Builder()
                    .borderTop(style.BorderTop())
                    .borderBottom(style.BorderBottom())
                    .borderLeft(style.BorderLeft())
                    .borderRight(style.BorderRight())
                    .topBorderColor(style.TopBorderColor())
                    .bottomBorderColor(style.BottomBorderColor())
                    .leftBorderColor(style.LeftBorderColor())
                    .rightBorderColor(style.RightBorderColor())
                    .alignment(style.Alignment())
                    .verticalAlignment(style.VerticalAlignment())
                    .fillBackgroundColor(style.FillBackgroundColor())
                    .dataFormat(style.DataFormat())
                    .forceCover(style.ForceCover())
                    .build();
        }else if (style != null && style.Type().equals(CONTENT_TYPE.ALL) && !CONTENT_TYPE.ALL.equals(type))
            return null;

        CellStyles styles = ( obj instanceof Class ) ?
                (CellStyles) ((Class)obj).getAnnotation(CellStyles.class) :
                ( obj instanceof Field ) ?
                        ((Field)obj).getAnnotation(CellStyles.class) :
                        null
                ;

        if ( styles != null ){
            CellStyle[] cellStyles = styles.value();
            for ( CellStyle cellStyle : cellStyles ){
                if (cellStyle.Type().equals(type)) {
                    style = cellStyle;
                    break;
                }
            }
        }else {
            if ( style == null || !style.Type().equals(type) )
                style = null;
        }

        return
                (style == null) ?
                        null :
                        new CellStyleProperties.Builder()
                                .borderTop(style.BorderTop())
                                .borderBottom(style.BorderBottom())
                                .borderLeft(style.BorderLeft())
                                .borderRight(style.BorderRight())
                                .topBorderColor(style.TopBorderColor())
                                .bottomBorderColor(style.BottomBorderColor())
                                .leftBorderColor(style.LeftBorderColor())
                                .rightBorderColor(style.RightBorderColor())
                                .alignment(style.Alignment())
                                .verticalAlignment(style.VerticalAlignment())
                                .fillBackgroundColor(style.FillBackgroundColor())
                                .dataFormat(style.DataFormat())
                                .forceCover(style.ForceCover())
                                .build();
    }

}