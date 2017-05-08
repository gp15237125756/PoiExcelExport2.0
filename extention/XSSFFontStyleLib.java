package com.ld.datacenter.poi.extention;

import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ld.datacenter.poi.convert.FontStyleConvert;
import com.ld.datacenter.poi.convert.IFontStyleConvert;
import com.ld.datacenter.poi.exception.FontStyleNotFoundException;
import com.ld.datacenter.poi.font.ColorCellFontImpl;
import com.ld.datacenter.poi.font.DefaultCellFontImpl;
<<<<<<< HEAD
import com.ld.datacenter.poi.font.TitleCellFontImpl;
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
import com.ld.datacenter.poi.helper.FontStyle;

public class XSSFFontStyleLib {

	/** 樣式緩存*/
	private  final Map<FontStyle,XSSFFont> FONT_STYLE_MAP=new LinkedHashMap<FontStyle, XSSFFont>();
	
<<<<<<< HEAD
	@SuppressWarnings("unused")
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	private  final XSSFWorkbook wb;
	 
	public XSSFFontStyleLib(XSSFWorkbook wb){
		this.wb=wb;
		IFontStyleConvert convert=new FontStyleConvert();
		//添加cell樣式 
		convert.convert(FontStyle.NONE, new DefaultCellFontImpl(wb),this);
		convert.convert(FontStyle.DEFAULT, new DefaultCellFontImpl(wb),this);
<<<<<<< HEAD
		convert.convert(FontStyle.TITLE, new TitleCellFontImpl(wb),this);
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
		convert.convert(FontStyle.COLOR, new ColorCellFontImpl(wb),this);
	}
	
	public  void addStyle(FontStyle k,XSSFFont v){
		FONT_STYLE_MAP.put(k, v);
	}
	
	public  void removeStyle(FontStyle k){
		FONT_STYLE_MAP.remove(k);
	}
	
	public  XSSFFont get(FontStyle k){
		XSSFFont v=FONT_STYLE_MAP.get(k);
		if(null==v){
			throw new FontStyleNotFoundException("the font style have not been defined!");
		}
		return v;
	}
	
	public  void clear(){
		FONT_STYLE_MAP.clear();
	}
}
