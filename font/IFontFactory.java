package com.ld.datacenter.poi.font;

import org.apache.poi.xssf.usermodel.XSSFFont;
/**
 * font factory
 * @author Cruz
 * @Date 2017/3/31
 * @version 01-00
 */
public interface IFontFactory {
	
	/**
	 * @param wb
	 * @return {@code XSSFFont}
	 */
	XSSFFont create();
}
