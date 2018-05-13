/*
 *  Created by ZhongWenjie on 2018-05-07 12:57
 */

package com.excel.poi.properties;

import org.apache.poi.ss.usermodel.CellType;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class EntityProperties {

    private Map<String,CellType> cellTypeMap = new ConcurrentHashMap<String,CellType>();

    public boolean put(String key, CellType cellType){
        cellTypeMap.put(key,cellType);
        return true;
    }

    public CellType get(String key){
        return cellTypeMap.get(key);
    }


}
