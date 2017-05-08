package com.ld.datacenter.poi.util;

import java.awt.Color;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.BorderStyle;
<<<<<<< HEAD
import org.apache.poi.ss.usermodel.CellStyle;
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.Assert;
/**
 * {@code XSSFWorkbook} excel utility,simplify export procedure
 * @author Cruz
 * @since  01-00
 */
<<<<<<< HEAD
@SuppressWarnings("deprecation")
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
public class XSSFCellUtil {
	//create cell
	public static XSSFCell createCellIfNotPresent(XSSFRow row,int colIndex){
		String message="XSSFRow must not be null!";
		Objects.requireNonNull(row, () -> message);
		XSSFCell cell=row.getCell(colIndex);
		if(Objects.isNull(cell)){
			cell=row.createCell(colIndex);
		}
		return cell;
	}
	
<<<<<<< HEAD
=======
	@SuppressWarnings("deprecation")
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	public static void createMergedRegionIfNotPresent(XSSFSheet sheet,int firstRow, int lastRow, int firstCol, int lastCol){
		String message="XSSFSheet must not be null!";
		Objects.requireNonNull(sheet, () -> message);
		sheet.addMergedRegion(new CellRangeAddress(firstRow,  lastRow,  firstCol,  lastCol)); 
	}
	
	
	
	//default cell style
	public static XSSFCellStyle createXSSFCellStyle(XSSFWorkbook wb){
		String message="XSSFWorkbook must not be null!";
		Objects.requireNonNull(wb, () -> message);
		XSSFCellStyle cellStyle=wb.createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		return cellStyle;
	}
	//vertical and horizon center aligned
 	public static XSSFCellStyle createCenterXSSFCellStyle(XSSFWorkbook wb){
		String message="XSSFWorkbook must not be null!";
		Objects.requireNonNull(wb, () -> message);
		XSSFCellStyle cellStyle=wb.createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		return cellStyle;
	}
<<<<<<< HEAD
 	
 	//vertical and horizon center aligned
 	 	public static XSSFCellStyle createTitleXSSFCellStyle(XSSFWorkbook wb){
 			String message="XSSFWorkbook must not be null!";
 			Objects.requireNonNull(wb, () -> message);
 			XSSFCellStyle cellStyle=wb.createCellStyle();
 			cellStyle.setWrapText(true);
 			cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
 			cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
 			cellStyle.setBorderBottom(BorderStyle.THIN);
 			cellStyle.setBorderLeft(BorderStyle.THIN);
 			cellStyle.setBorderTop(BorderStyle.THIN);
 			cellStyle.setBorderRight(BorderStyle.THIN);
 			cellStyle.setFillForegroundColor(new XSSFColor( new Color(75, 172, 198)));
 			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
 			return cellStyle;
 		}
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	
	
	/**
	 * @param wb
	 * @param color
	 * @param foreGround
	 * @return
	 */
	public static XSSFCellStyle createBackgroundColorXSSFCellStyle(XSSFWorkbook wb,XSSFColor color,short foreGround){
		String message="XSSFWorkbook must not be null!";
		Objects.requireNonNull(wb, () -> message);
		XSSFCellStyle cellStyle=wb.createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFillForegroundColor(color);
		cellStyle.setFillPattern(foreGround);
		return cellStyle;
	}
	
	/**
	 * @param cell
	 * @param value
	 * @param style
	 */
	public static void setCellProperties(XSSFCell cell,String value,XSSFCellStyle style){
		Assert.notNull(cell, "cell must not be null!");
		//fill year
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}
	
	/**
	 * 返回副标题字体样式
	 * @param wb
	 * @return
	 */
	public static XSSFCellStyle getCaptionCellStyle(XSSFWorkbook wb){
		//default cellStyle
		XSSFCellStyle cellStyle=createXSSFCellStyle(wb);
		//default font
		XSSFFont cellfont=XSSFontUtil.createXSSFFont(wb);
		cellStyle.setFont(cellfont);
		return cellStyle;
	}
	
	/**
	 * 返回excel内容字体样式
	 * @param wb
	 * @return
	 */
	public static XSSFCellStyle getcontentCellStyle(XSSFWorkbook wb){
		String message="XSSFWorkbook must not be null!";
		Objects.requireNonNull(wb, () -> message);
		XSSFCellStyle contentStyle=createCenterXSSFCellStyle(wb);
		//fill enlarged font
		XSSFFont enlargedFont=XSSFontUtil.createColorXSSFFont(wb, "新宋体", (short) 16, new XSSFColor(Color.WHITE), true);
		contentStyle.setFont(enlargedFont);
		return contentStyle;
	}
	
	
	/**
	 * 设置单行 某列内容
	 * @param row
	 * @param col
	 * @param value
	 * @param wb
	 */
	public static void setRowContents(XSSFRow row,int col,String value,XSSFWorkbook wb){
		String message="XSSFWorkbook must not be null!";
		Objects.requireNonNull(wb, () -> message);
		XSSFCell cell=createCellIfNotPresent(row, col);
		setCellProperties(cell, value, getcontentCellStyle(wb));
	}
	
	
	/**
	 * 设置单行全部内容
	 * @param sheet
	 * @param wb
	 * @param rowNum
	 * @param kVal
	 */
	public static void setRowMutipleColContent(XSSFSheet sheet,XSSFWorkbook wb,int rowNum,final Map<Integer,String> kVal){
		String message="XSSFSheet must not be null!";
		Objects.requireNonNull(sheet, () -> message);
		XSSFRow row=sheet.getRow(rowNum);
		if(!kVal.isEmpty()){
			kVal.forEach((k,v)->{
				setRowContents(row,k,v,wb);
			});
		}
		
	}
	
	public static void insertDedicatedStyleCell(XSSFRow row,int col,XSSFCellStyle style,String value){
		String message="XSSFRow or XSSFCellStyle must not be null!";
		Objects.requireNonNull(row, () -> message);
		Objects.requireNonNull(style, () -> message);
<<<<<<< HEAD
=======
		XSSFCell cell=createCellIfNotPresent(row,col);
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
		setCellProperties(createCellIfNotPresent(row, col), value, style);
	}
	
	
}
