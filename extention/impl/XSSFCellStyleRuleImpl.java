package com.ld.datacenter.poi.extention.impl;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import com.ld.datacenter.poi.extention.XSSFCellStyleLib;
import com.ld.datacenter.poi.extention.XSSFCellStyleRule;
import com.ld.datacenter.poi.helper.CellStyle;

public class XSSFCellStyleRuleImpl implements XSSFCellStyleRule{

	@Override
	public void registerStyleRule(CellStyle k, XSSFCellStyle v,
			XSSFCellStyleLib lib) {
		
		lib.addStyle(k, v);
	}

	@Override
	public void removeStyleRule(CellStyle k, XSSFCellStyleLib lib) {
		lib.removeStyle(k);
		
	}

	

}
