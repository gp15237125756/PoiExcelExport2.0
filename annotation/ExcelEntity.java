package com.ld.datacenter.poi.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.ld.datacenter.poi.helper.CellStyle;
import com.ld.datacenter.poi.helper.FontStyle;

/**
 * 标记实体为Excel实体
 */
@Target({ ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelEntity {
	/**
	 * 年度位置
	 */
	int[] yrPos() default 0;
<<<<<<< HEAD
	/**
	 * 
	 * 年度格子样式
	 */
	CellStyle yrCellStyle() default CellStyle.NONE;
	/**
	 * 年度格子字体
	 */
	FontStyle yrFontStyle() default FontStyle.NONE;
=======
	

	CellStyle yrCellStyle() default CellStyle.NONE;
	
	FontStyle yrFontStyle() default FontStyle.NONE;

>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	/**
	 * 季度位置
	 */
	int[] quaPos() default 0;
<<<<<<< HEAD
	/**
	 * 季度格子样式
	 */
=======
	
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	CellStyle quaCellStyle() default CellStyle.NONE;
	
	FontStyle quaFontStyle() default FontStyle.NONE;

	/**
	 * 月份位置
	 */
	int[] monPos() default 0;
<<<<<<< HEAD
	/**
	 * 月份格子样式
	 */
	CellStyle monCellStyle() default CellStyle.NONE;
	
	FontStyle monFontStyle() default FontStyle.NONE;
	/**
	 * 行号
	 */
	int row() default 0;
	/**
	 * 列号
	 */
	int col() default 0;
=======
	
	CellStyle monCellStyle() default CellStyle.NONE;
	
	FontStyle monFontStyle() default FontStyle.NONE;
	
	int row() default 0;
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	
	
}
