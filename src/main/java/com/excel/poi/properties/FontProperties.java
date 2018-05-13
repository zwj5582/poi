/*
 *  Created by ZhongWenjie on 2018-05-07 16:27
 */

package com.excel.poi.properties;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.hssf.util.HSSFColor;

@Getter@Setter
@NoArgsConstructor@AllArgsConstructor
public class FontProperties {

    private String FontName = "宋体";

    private boolean Bold = false;

    private HSSFColor.HSSFColorPredefined Color = HSSFColor.HSSFColorPredefined.BLACK;

    private short FontHeightInPoints = 12;

    private boolean forceCover = false;

}
