/*
 *  Created by ZhongWenjie on 2018-05-07 20:33
 */

package com.excel.poi.support;

import ognl.Ognl;
import ognl.OgnlException;

import java.util.*;

public class MappingObjectProxy {

    private Object target;

    private List<Object> valueList = new ArrayList<Object>();

    private boolean OGNL = false;

    public MappingObjectProxy(Object target) {
        if (target instanceof Iterable)
            for (Object cellValue : ((Iterable) target))
                valueList.add(cellValue);
        else if ( target.getClass().isArray() )
            valueList.addAll(Arrays.asList( (Object[])target) );
        else{
            OGNL = true;
        }
        this.target = target;
    }

    public List<Object> getList( List<String> expressions ) throws OgnlException {
        if (!OGNL) return valueList;
        for ( String exp : expressions )
            valueList.add(Ognl.getValue(exp,target));
        return valueList;
    }

}
