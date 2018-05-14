/*
 *  Created by ZhongWenjie on 2018-05-11 11:02
 */

package com.excel.poi.chooser;

import com.excel.poi.annotation.ConditionalStyle;
import com.excel.poi.annotation.DEFAULT;

import java.util.List;

public class CellStyleChooser {

    public static final ConditionalChooser ALWAYS_TRUE = new ConditionalChooser() {
        @Override
        public boolean doConditionalChoose(List<Object> valueList, Object input, int row) {
            return true;
        }
    };

    public static final ConditionalStyle EMPTY = DEFAULT.class.getAnnotation(ConditionalStyle.class);

    private ConditionalStyle TRUE;

    private ConditionalStyle FALSE;

    private ConditionalChooser chooser;

    public CellStyleChooser(ConditionalChooser chooser, ConditionalStyle TRUE, ConditionalStyle FALSE) {
        this.TRUE = TRUE;
        this.FALSE = FALSE;
        this.chooser = chooser;
    }

    public ConditionalStyle fetchCellStyle(List<Object> valueList,Object input, int row){
        return chooser.doConditionalChoose(valueList,input,row) ? TRUE : FALSE;
    }

}
