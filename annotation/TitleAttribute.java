package com.ld.datacenter.poi.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Repeatable;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.ld.datacenter.poi.helper.CellStyle;
import com.ld.datacenter.poi.helper.FontStyle;

/**
 * 标记excel标题栏
 */
@Repeatable(TitleAttributes.class)
@Target({ ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface TitleAttribute {
	
	/**
	 * 格子样式
	 */
	CellStyle cellStyle() default CellStyle.NONE;
	/**
	 * 格子字体
	 */
	FontStyle fontStyle() default FontStyle.NONE;
	/**
	 * 行号
	 */
	int row() default 0;
	/**
	 * 列号
	 */
	int col() default 0;
	
	
}
