package com.ld.datacenter.poi.extention;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ld.datacenter.poi.convert.CellStyleConvert;
import com.ld.datacenter.poi.convert.ICellStyleConvert;
import com.ld.datacenter.poi.exception.CellStyleNotFoundException;
import com.ld.datacenter.poi.helper.CellStyle;
import com.ld.datacenter.poi.style.DefaultCellStyleImpl;
import com.ld.datacenter.poi.style.TextCellStyleImpl;
<<<<<<< HEAD
import com.ld.datacenter.poi.style.TitleCellStyleImpl;
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
/**
 * 定義cell樣式庫，可以動態追加
 * @author Cruz
 * @version 01-00
 * @Date 2017/3/31 13:22
 */
public class XSSFCellStyleLib {
	/** 樣式緩存*/
	private  final Map<CellStyle,XSSFCellStyle> CELL_STYLE_MAP=new LinkedHashMap<CellStyle, XSSFCellStyle>();
	
<<<<<<< HEAD
	@SuppressWarnings("unused")
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	private  final XSSFWorkbook wb;
	 
	public XSSFCellStyleLib(XSSFWorkbook wb){
		this.wb=wb;
		ICellStyleConvert convert=new CellStyleConvert();
		//添加cell樣式 
		convert.convert(CellStyle.NONE, new DefaultCellStyleImpl(wb),this);
		convert.convert(CellStyle.DEFAULT, new DefaultCellStyleImpl(wb),this);
		convert.convert(CellStyle.STANDARD, new TextCellStyleImpl(wb),this);
<<<<<<< HEAD
		convert.convert(CellStyle.TITLE, new TitleCellStyleImpl(wb),this);
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	}
	
	
	public  void addStyle(CellStyle k,XSSFCellStyle v){
		CELL_STYLE_MAP.put(k, v);
	}
	
	public  void removeStyle(CellStyle k){
		CELL_STYLE_MAP.remove(k);
	}
	
	public  XSSFCellStyle get(CellStyle k){
		XSSFCellStyle v=CELL_STYLE_MAP.get(k);
		if(null==v){
			throw new CellStyleNotFoundException("the cell style have not been defined!");
		}
		return v;
	}
	
	public  void clear(){
		CELL_STYLE_MAP.clear();
	}
}
