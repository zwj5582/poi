/*
 *  Created by ZhongWenjie on 2018-05-07 17:21
 */

package com.excel.poi.context;

import com.excel.poi.chooser.CellStyleChooser;
import com.excel.poi.render.CellStyleReRender;
import com.excel.poi.properties.CellStyleProperties;
import com.excel.poi.properties.FontProperties;
import lombok.Getter;
import lombok.Setter;

@Getter@Setter
public class PerCellStyleHolder {

    private String title;

    private CellStyleProperties perHeaderCellStyleProperties;

    private CellStyleProperties perBodyCellStyleProperties;

    private FontProperties bodyFontProperties;

    private FontProperties headFontProperties;

    private CellStyleReRender render;

    private CellStyleChooser chooser;

}
