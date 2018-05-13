/*
 *  Created by ZhongWenjie on 2018-05-08 15:56
 */

package com.excel.poi.exception;

public class TooManyDataExportException extends Exception {

    public TooManyDataExportException(int length) {
        super("Too Many Data Export: length ["+length+"]");
    }
}
