package com.ld.datacenter.poi.style;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
/**
 * create cell style 
 * @author Cruz
 * @version 01-00
 * @Date 2017/3/31
 * 
 */
public interface ICellStyleFactory {
	
	/**
	 * @param wb
	 * @return {@code XSSFCellStyle}
	 */
	XSSFCellStyle create();
}
