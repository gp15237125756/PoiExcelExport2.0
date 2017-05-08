package com.ld.datacenter.poi.convert;

import com.ld.datacenter.poi.extention.XSSFCellStyleLib;
import com.ld.datacenter.poi.helper.CellStyle;
import com.ld.datacenter.poi.style.ICellStyleFactory;

public interface ICellStyleConvert {

	void convert(CellStyle style,ICellStyleFactory fac,XSSFCellStyleLib lib);
}
