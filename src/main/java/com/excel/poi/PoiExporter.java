/*
 *  Created by ZhongWenjie on 2018-05-08 17:05
 */

package com.excel.poi;

import com.excel.poi.annotation.ConditionalStyle;
import com.excel.poi.context.PoiContext;
import com.excel.poi.dataSource.PoiPageable;
import com.excel.poi.dataSource.impl.ComposePageableDataSource;
import com.excel.poi.dataSource.impl.PoiDataSource;
import com.excel.poi.exception.TooManyDataExportException;
import com.excel.poi.context.PerCellStyleHolder;
import com.excel.poi.reader.DefaultPoiPropertiesReader;
import com.excel.poi.reader.PoiPropertiesReader;
import com.excel.poi.render.Render;
import com.excel.poi.support.CONTENT_TYPE;
import com.excel.poi.support.MappingObjectProxy;
import lombok.Getter;
import ognl.OgnlException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.util.*;

public class PoiExporter {

    private PoiPropertiesReader propertiesReader;

    private @Getter HSSFWorkbook workbook;

    private @Getter PoiContext poiContext;

    public PoiExporter() {
        this.propertiesReader = new DefaultPoiPropertiesReader();
    }

    public PoiExporter(PoiPropertiesReader propertiesReader) {
        this.propertiesReader = propertiesReader;
    }

    @SuppressWarnings("unchecked")
    public HSSFWorkbook createExporter(Object entity,Object dataSource)
            throws TooManyDataExportException, OgnlException, IllegalAccessException, InstantiationException, ClassNotFoundException {

        this.workbook = new HSSFWorkbook();

        Class entityClass = entity.getClass();

        this.poiContext = this.propertiesReader.readProperties(entityClass);
        poiContext.setWorkbook(workbook);

        Render render = new Render(this);

        PoiPageable page = new ComposePageableDataSource(dataSource);

        Iterable<PoiDataSource> poiDataSources = page.transformDataSource();

        int sheetNum = 1;
        for ( PoiDataSource source : poiDataSources ) {

            String title = fetchTitle();

            HSSFSheet sheet = workbook.createSheet(title);

            List<String> titleNames = fetchTitleNames();
            int columnHeight = fetchGlobalColumnHeight();
            int columnWidth = fetchGlobalColumnWidth();

            int col_index = render.renderGlobal(sheet, sheetNum++, columnWidth, title, titleNames);

            col_index = render.renderHeader(sheet, columnHeight, titleNames, col_index);

            Iterable<MappingObjectProxy> iterable = source.fetchDataSource();

            render.renderBody(sheet,columnHeight,columnWidth,col_index,iterable);
        }
        return workbook;
    }

    public void fillValue(HSSFCell cell, CellType cellType, Object val) {
        switch (cellType) {
            case STRING:
                if (val != null)
                    cell.setCellValue(val.toString());
                break;
            case BLANK:
                cell.setCellValue("");
                break;
            case BOOLEAN:
                if (val != null)
                    cell.setCellValue((Boolean)val);
                break;
            case NUMERIC:
                if ( val != null && val instanceof Date){
                    cell.setCellValue((Date)val);
                } else if ( val !=null )
                    cell.setCellValue(Double.valueOf(val.toString()));
                break;
            case _NONE:
                break;
        }
    }

    public CellType fetchCellType(int index) {
        return poiContext.fetchCellType(index);
    }

    public HSSFCellStyle createCellStyle(CONTENT_TYPE type, int index) {
        return poiContext.createCellStyle(type,index);
    }

    public ConditionalStyle fetchConditionalStyle(PerCellStyleHolder perCellStyleHolder, List<Object> valueList, Object val, HSSFCellStyle cellStyle, int colIndex) {
        return poiContext.fetchConditionalStyle(valueList, val, cellStyle, perCellStyleHolder,colIndex);
    }

    public PerCellStyleHolder fetchPerCellStyleHolder(int index) {
        return poiContext.fetchPerCellStyleHolder(index);
    }

    public List<String> fetchExpressions() {
        return poiContext.getExpressions();
    }

    public int fetchTitleColumnHeight(){
        return poiContext.getGlobalProperties().getTitleColumnHeight();
    }

    private int fetchGlobalColumnHeight() {
        return poiContext.getGlobalProperties().getColumnHeight();
    }

    private int fetchGlobalColumnWidth() {
        return poiContext.getGlobalProperties().getColumnWidth();
    }

    private List<String> fetchTitleNames() {
        return poiContext.getGlobalProperties().getTitleNames();
    }

    private String fetchTitle() {
        return poiContext.getGlobalProperties().getTitle();
    }



}
