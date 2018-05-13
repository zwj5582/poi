/*
 *  Created by ZhongWenjie on 2018-05-07 21:40
 */

package com.excel.poi.dataSource.impl;


import com.excel.poi.support.MappingObjectProxy;

import java.util.ArrayList;
import java.util.List;

public class ArrayDataSource extends PoiDataSource<Object[]> {

    public ArrayDataSource(Object[] dataSource) {
        super(dataSource);
    }

    @Override
    public Iterable<MappingObjectProxy> fetchDataSource() {
        return this.dataSources;
    }

    protected ArrayDataSource(Iterable<MappingObjectProxy> dataSources, int length) {
        super(dataSources, length);
    }

    @Override
    protected void initPoiDataSource() {
        List<MappingObjectProxy> source = new ArrayList<MappingObjectProxy>();
        for (Object row :  dataSource)
            source.add(new MappingObjectProxy(row));
        super.length = (short) source.size();
        this.dataSources = source;
    }

    @Override
    public int length() {
        return super.length;
    }
}