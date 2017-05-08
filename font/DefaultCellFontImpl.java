package com.ld.datacenter.poi.font;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ld.datacenter.poi.util.XSSFontUtil;

public class DefaultCellFontImpl implements IFontFactory {
	private XSSFWorkbook wb;
	
	public DefaultCellFontImpl(XSSFWorkbook wb){
		this.wb=wb;
	}
	@Override
	public XSSFFont create() {
		// TODO Auto-generated method stub
		return XSSFontUtil.createXSSFFont(this.wb);
	}

}
