package com.ld.datacenter.poi.font;

import java.awt.Color;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ld.datacenter.poi.util.XSSFontUtil;

public class ColorCellFontImpl implements IFontFactory {
	private XSSFWorkbook wb;
	
	public ColorCellFontImpl(XSSFWorkbook wb){
		this.wb=wb;
	}
	@Override
	public XSSFFont create() {
		// TODO Auto-generated method stub
		return XSSFontUtil.createColorXSSFFont(this.wb, "新宋体", (short) 16, new XSSFColor(Color.WHITE), true);
	}

}
