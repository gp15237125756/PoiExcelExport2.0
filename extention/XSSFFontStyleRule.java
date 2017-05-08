package com.ld.datacenter.poi.extention;

import org.apache.poi.xssf.usermodel.XSSFFont;

import com.ld.datacenter.poi.helper.FontStyle;

public interface XSSFFontStyleRule {
	/**
	 * 註冊字體
	 */
	void registerStyleRule(FontStyle k,XSSFFont v,XSSFFontStyleLib lib);
	/**
	 * 移除字體
	 */
	void removeStyleRule(FontStyle k,XSSFFontStyleLib lib);
}
