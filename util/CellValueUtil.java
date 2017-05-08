package com.ld.datacenter.poi.util;

import java.util.Objects;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.Assert;

import com.ld.datacenter.poi.extention.XSSFCellStyleLib;
import com.ld.datacenter.poi.extention.XSSFFontStyleLib;
import com.ld.datacenter.poi.helper.CellStyle;
import com.ld.datacenter.poi.helper.FontStyle;
/**
 * template utility
 * @author Cruz
 * @version 01-00
 * @Date 2017/3/31
 */
public class CellValueUtil {
	// create cell
	public static XSSFCell createCellIfNotPresent(XSSFRow row, int colIndex) {
		String message = "XSSFRow must not be null!";
		Objects.requireNonNull(row, () -> message);
		XSSFCell cell = row.getCell(colIndex);
		if (Objects.isNull(cell)) {
			cell = row.createCell(colIndex);
		}
		return cell;
	}

	/**
	 * @param cell
	 * @param value
	 * @param style
	 */
	public static void setCellProperties(XSSFCell cell, String value,
			XSSFCellStyle style) {
		Assert.notNull(cell, "cell must not be null!");
		// fill year
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}
	
	/**
	 * @param sheet 
	 * @param pos 位置【row,col】
	 * @param style
	 * @param v
	 */
	public static void setCell(final XSSFWorkbook wb,final XSSFSheet sheet,int[] pos,CellStyle style,FontStyle font,String v){
		XSSFRow row = sheet.getRow(pos[0]);
		XSSFCell cell = CellValueUtil.createCellIfNotPresent(row, pos[1]);
		XSSFCellStyle _style = Objects
				.requireNonNull(new XSSFCellStyleLib(wb).get(style),
						"specified style not defined!");
		XSSFFont _font = Objects
				.requireNonNull(new XSSFFontStyleLib(wb).get(font),
						"specified font not defined!");
		_style.setFont(_font);
		setCellProperties(cell, v, _style);
	}
	
	
}
