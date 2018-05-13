/*
 *  Created by ZhongWenjie on 2018-05-07 11:29
 */

package com.excel.poi.properties;

import lombok.Getter;
import lombok.Setter;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

@Getter@Setter
public class GlobalProperties {

    private String title;

    private int columnWidth;

    private int columnHeight;

    private int titleColumnHeight;

    private List<String> titleNames = new ArrayList<String>();

    private CellStyleProperties titleCellStyleProperties;

    private FontProperties titleFontProperties;

    public boolean put(String key, Object value){
        Class<? extends GlobalProperties> selfClass = this.getClass();
        Field field = null;
        try {
            field = selfClass.getDeclaredField(key);
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        }
        if (field == null) return false;
        field.setAccessible(true);
        try {
            field.set(this,value);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    public String getTitleNamesByIndex(int index) {
        if (titleNames  == null || titleNames.isEmpty() || index < 0 || index > titleNames.size() - 1 )
            return PropertiesConstants.EMPTY;
        String titleName = titleNames.get(index);
        if ( titleName == null ) return PropertiesConstants.EMPTY;
        return titleName;
    }
}
