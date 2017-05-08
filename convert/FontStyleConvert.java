package com.ld.datacenter.poi.convert;

import com.ld.datacenter.poi.extention.XSSFFontStyleLib;
import com.ld.datacenter.poi.extention.XSSFFontStyleRule;
import com.ld.datacenter.poi.extention.impl.XSSFFontStyleRuleImpl;
import com.ld.datacenter.poi.font.IFontFactory;
import com.ld.datacenter.poi.helper.FontStyle;

public class FontStyleConvert implements IFontStyleConvert {

	@Override
	public void convert(FontStyle style, IFontFactory fac,XSSFFontStyleLib lib) {
		XSSFFontStyleRule rule=new XSSFFontStyleRuleImpl();
		rule.registerStyleRule(style, fac.create(),lib);
	}

}
