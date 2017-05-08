package com.ld.datacenter.poi.style;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ld.datacenter.util.XSSFCellUtil;
/**
 * caption cell
 */
public class DefaultCellStyleImpl implements ICellStyleFactory {
	
	private XSSFWorkbook wb;
	
	public DefaultCellStyleImpl(XSSFWorkbook wb){
		this.wb=wb;
	}
	
	@Override
	public XSSFCellStyle create() {
		return XSSFCellUtil.createXSSFCellStyle(this.wb);
	}

}
