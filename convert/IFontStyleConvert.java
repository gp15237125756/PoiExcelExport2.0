package com.ld.datacenter.poi.convert;

import com.ld.datacenter.poi.extention.XSSFFontStyleLib;
import com.ld.datacenter.poi.font.IFontFactory;
import com.ld.datacenter.poi.helper.FontStyle;

public interface IFontStyleConvert {
	void convert(FontStyle style,IFontFactory fac,XSSFFontStyleLib lib);
}
