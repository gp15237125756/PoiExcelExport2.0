package com.ld.datacenter.poi.extention;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import com.ld.datacenter.poi.helper.CellStyle;

public interface XSSFCellStyleRule {
	/**
	 * 註冊樣式
	 */
	void registerStyleRule(CellStyle k,XSSFCellStyle v,XSSFCellStyleLib lib);
	/**
	 * 移除樣式
	 */
	void removeStyleRule(CellStyle k,XSSFCellStyleLib lib);
	
}
