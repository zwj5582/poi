/*
 *  Created by ZhongWenjie on 2018-05-13 20:30
 */

package com.excel.poi.render;

import com.excel.poi.PoiExporter;
import com.excel.poi.annotation.ConditionalStyle;
import com.excel.poi.context.PerCellStyleHolder;
import com.excel.poi.support.CONTENT_TYPE;
import com.excel.poi.render.CellStyleReRender;
import com.excel.poi.support.MappingObjectProxy;
import ognl.OgnlException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class Render {

    private PoiExporter poiExporter;

    public Render(PoiExporter poiExporter) {
        this.poiExporter = poiExporter;
    }

    public int renderGlobal(HSSFSheet sheet, int sheetNum, int columnWidth, String title, List<String> titleNames) {

        int titleColumnHeight = poiExporter.fetchTitleColumnHeight();

        sheet.setColumnWidth(0, columnWidth);

        if (sheetNum != 1 ) return 0;

        sheet.addMergedRegion(new CellRangeAddress(0,(short)0,0,(short)titleNames.size()-1));
        HSSFRow titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(titleColumnHeight);
        HSSFCell titleRowCell = titleRow.createCell(0);
        titleRowCell.setCellValue(title);
        HSSFCellStyle cellStyle = poiExporter.createCellStyle(CONTENT_TYPE.TITLE, 0);

        titleRowCell.setCellStyle(cellStyle);

        for (int i = 1; i < titleNames.size(); i++) {
            HSSFCell cell = titleRow.createCell(i);
            cell.setCellStyle(cellStyle);
        }

        return 1;
    }

    public int renderHeader(HSSFSheet sheet, int columnHeight, List<String> titleNames, int col_index) {
        HSSFRow header = sheet.createRow(col_index++);
        header.setHeightInPoints(columnHeight);
        for (int index = 0; index < titleNames.size(); index++) {
            HSSFCell cell = header.createCell(index);
            HSSFCellStyle cellStyle =poiExporter.createCellStyle(CONTENT_TYPE.HEADER, index);
            cell.setCellValue(titleNames.get(index));
            cell.setCellStyle(cellStyle);
        }
        return col_index;
    }

    public void renderBody(HSSFSheet sheet, int columnHeight, int columnWidth, int col_index, Iterable<MappingObjectProxy> iterable) throws OgnlException {
        for (MappingObjectProxy proxy : iterable) {
            sheet.setColumnWidth(col_index, columnWidth);
            HSSFRow row = sheet.createRow(col_index++);

            row.setHeightInPoints(columnHeight);

            List<ConditionalStyle> styleList = new ArrayList<ConditionalStyle>();
            List<Object> valueList = proxy.getList(poiExporter.fetchExpressions());

            renderRow(row,valueList,styleList,col_index);

            reRenderRow(row,valueList,styleList,col_index);
        }
    }

    private void renderRow(HSSFRow row, List<Object> valueList, List<ConditionalStyle> styleList, int col_index) {
        for (int i = 0; i < valueList.size(); i++) {
            Object val = valueList.get(i);

            HSSFCellStyle cellStyle = poiExporter.createCellStyle(CONTENT_TYPE.BODY, i);
            PerCellStyleHolder perCellStyleHolder = poiExporter.fetchPerCellStyleHolder(i);
            CellType cellType = poiExporter.fetchCellType(i);
            HSSFCell cell = row.createCell(i, cellType);

            ConditionalStyle conditionalStyle =
                    poiExporter.fetchConditionalStyle(perCellStyleHolder,valueList, val,cellStyle, col_index-1);
            styleList.add(conditionalStyle);

            cell.setCellStyle(cellStyle);
            poiExporter.fillValue(cell, cellType,val);
        }
    }

    private void reRenderRow(HSSFRow row, List<Object> valueList, List<ConditionalStyle> styleList, int col_index) {
        CellStyleReRender rowCellStyleReRender = new CellStyleReRender(styleList);

        for (int i = 0; i < valueList.size(); i++) {
            HSSFCell cell = row.getCell(i);
            HSSFCellStyle cellStyle = cell.getCellStyle();

            cellStyle = rowCellStyleReRender.reRender(poiExporter.getWorkbook(), cellStyle, valueList,null, col_index - 1);

            cell.setCellStyle(cellStyle);
        }
    }

}
