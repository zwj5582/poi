/*
 *  Created by ZhongWenjie on 2018-05-07 17:49
 */

package com.excel.poi.dataSource.impl;

import com.excel.poi.support.MappingObjectProxy;

import java.util.ArrayList;
import java.util.List;

public class IterableDataSource extends PoiDataSource<Iterable> {

    public IterableDataSource(Iterable dataSource) {
        super(dataSource);
    }

    @Override
    public Iterable<MappingObjectProxy> fetchDataSource() {
        return super.dataSources;
    }

    protected IterableDataSource(Iterable<MappingObjectProxy> dataSources, int length) {
        super(dataSources, length);
    }

    @Override
    protected void initPoiDataSource() {
        List<MappingObjectProxy> source = new ArrayList<MappingObjectProxy>();
        for (Object row :  dataSource)
            source.add(new MappingObjectProxy(row));
        super.length = (short) source.size();
        super.dataSources = source;
    }

    @Override
    public int length() {
        return super.length;
    }
}
