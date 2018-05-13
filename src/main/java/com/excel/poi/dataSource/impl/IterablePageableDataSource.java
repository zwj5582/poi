/*
 *  Created by ZhongWenjie on 2018-05-08 10:15
 */

package com.excel.poi.dataSource.impl;

import com.excel.poi.support.MappingObjectProxy;
import com.excel.poi.exception.TooManyDataExportException;
import com.excel.poi.dataSource.PoiPageable;

import java.util.ArrayList;
import java.util.List;

public class IterablePageableDataSource implements PoiPageable {

    private List<PoiDataSource> pageSource = new ArrayList<PoiDataSource>();

    private short size = 1;

    private IterableDataSource iterableDataSource;

    public IterablePageableDataSource(Iterable dataSource) {
        this.iterableDataSource = new IterableDataSource(dataSource);
    }

    @Override
    public Iterable<PoiDataSource> transformDataSource() throws TooManyDataExportException {
        if ( iterableDataSource.length <= MAX_DEFAULT_PAGE_SIZE  ) {
            pageSource.add(iterableDataSource);
            return pageSource;
        }else if ( iterableDataSource.length >= (MAX_DEFAULT_PAGE_SIZE * MAX_PAGE_NUM) ){
            throw new TooManyDataExportException(iterableDataSource.length);
        }else {
            this.size = (short) (iterableDataSource.length % MAX_DEFAULT_PAGE_SIZE + 1 );
            int lastPageSize = iterableDataSource.length - ( MAX_DEFAULT_PAGE_SIZE * this.size );
            int index = 1;
            int current = 1;
            List<MappingObjectProxy> data = new ArrayList<MappingObjectProxy>();
            for (MappingObjectProxy proxy : iterableDataSource.dataSources ) {
                data.add(proxy);
                if ( index++ >= MAX_DEFAULT_PAGE_SIZE ) {
                    index = 1;
                    pageSource.add(new IterableDataSource(data, (current++ < this.size) ? MAX_DEFAULT_PAGE_SIZE : lastPageSize ));
                    data = new ArrayList<MappingObjectProxy>();
                }
            }
            return pageSource;
        }
    }

    @Override
    public short size() {
        return this.size;
    }

    @Override
    public short currentPage() {
        return 0;
    }
}
