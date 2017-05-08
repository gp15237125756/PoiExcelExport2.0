package com.ld.datacenter.poi.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import com.ld.datacenter.poi.helper.FontStyle;
import com.ld.datacenter.poi.helper.CellStyle;
/**
 * indicate field export available
 * @author Cruz
 * @version 01-00
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface EntityAttribute {
	/** excel column name*/
	String columnName() default "";
	/**	whether required*/
	boolean required() default false;

<<<<<<< HEAD
	int[] col() default 0;
=======
	int col() default 0;
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	
	int[] row() default 0;
	
	boolean merge() default false;
	
	CellStyle cellStyle() default CellStyle.NONE;
	
	FontStyle fontStyle() default FontStyle.NONE;
	
	boolean broken() default false;
}
