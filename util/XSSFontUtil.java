package com.ld.datacenter.poi.util;

import java.awt.Color;
import java.util.Objects;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFontUtil {
	//default
	/**
	 * @param wb
	 * @return {code XSSFFont}
	 */
	public static XSSFFont createXSSFFont(XSSFWorkbook wb){
		String message="XSSFWorkbook must not be null!";
		Objects.requireNonNull(wb, () -> message);
		XSSFFont cellfont=wb.createFont();
		cellfont.setFontName("新宋体");
		cellfont.setFontHeightInPoints((short) 12);
		cellfont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		cellfont.setColor(new XSSFColor(new Color(50,73,38)));
		
		return cellfont;
	}
	
	/**
	 * @param wb
	 * @param fontFamily
	 * @param height
	 * @param color
	 * @param bold
	 * @return
	 */
	public static XSSFFont createColorXSSFFont(XSSFWorkbook wb,String fontFamily,short height,XSSFColor color,boolean bold){
		String message="XSSFWorkbook must not be null!";
		Objects.requireNonNull(wb, () -> message);
		XSSFFont cellfont=wb.createFont();
		cellfont.setFontName(fontFamily);
		cellfont.setFontHeightInPoints(height);
		cellfont.setColor(color);
		if(bold){
			cellfont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		}
		return cellfont;
	}
<<<<<<< HEAD
	
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
}
