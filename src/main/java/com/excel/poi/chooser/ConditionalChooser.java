/*
 *  Created by ZhongWenjie on 2018-05-10 22:11
 */

package com.excel.poi.chooser;

import java.util.List;

public interface ConditionalChooser {

    boolean doConditionalChoose(List<Object> valueList , Object input, int row);

}
