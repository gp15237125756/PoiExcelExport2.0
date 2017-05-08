package com.ld.datacenter.poi.style;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ld.datacenter.poi.util.XSSFCellUtil;
/**
 * text cell
 * @author Cruz
 * @version 01-00
 */
public class TextCellStyleImpl implements ICellStyleFactory {
	
	private XSSFWorkbook wb;
	
	public TextCellStyleImpl(XSSFWorkbook wb){
		this.wb=wb;
	}
	@Override
	public XSSFCellStyle create() {
		return XSSFCellUtil.createCenterXSSFCellStyle(this.wb);
	}

}
