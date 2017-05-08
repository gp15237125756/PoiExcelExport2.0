package com.ld.datacenter.poi.extention.impl;

import org.apache.poi.xssf.usermodel.XSSFFont;

import com.ld.datacenter.poi.extention.XSSFFontStyleLib;
import com.ld.datacenter.poi.extention.XSSFFontStyleRule;
import com.ld.datacenter.poi.helper.FontStyle;

public class XSSFFontStyleRuleImpl implements XSSFFontStyleRule {

	@Override
	public void registerStyleRule(FontStyle k, XSSFFont v, XSSFFontStyleLib lib) {
		lib.addStyle(k, v);
	}

	@Override
	public void removeStyleRule(FontStyle k, XSSFFontStyleLib lib) {
		lib.removeStyle(k);
	}

	

	

	

}
