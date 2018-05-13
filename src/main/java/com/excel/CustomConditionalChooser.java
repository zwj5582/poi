/*
 *  Created by ZhongWenjie on 2018-05-10 22:54
 */

package com.excel;

import com.excel.poi.chooser.ConditionalChooser;

public class CustomConditionalChooser implements ConditionalChooser {

    @Override
    public boolean doConditionalChoose(Object input, int row) {
        return row%2 == 0 ;
    }

}
