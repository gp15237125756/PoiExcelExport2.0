package com.ld.datacenter.poi.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.ld.datacenter.poi.helper.RenderStyle;

/**
 * 标记数据填充渲染方向
 * （1）从左到右
 * （2）从上到下
 * 默认（1）方式
 *  类级别注解
 */
@Target({ ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface RenderDirection {
	/**
	 * 年度位置
	 */
	RenderStyle value() default RenderStyle.HORIZATION;	
	
}
