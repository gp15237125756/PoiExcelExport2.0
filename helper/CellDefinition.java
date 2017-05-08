package com.ld.datacenter.poi.helper;

import java.lang.reflect.Field;

import com.ld.datacenter.poi.annotation.EntityAttribute;

/**
 * 定义导出单个Cell的格式
 * 定义cell位置、样式、字体、是否合并单元格等属性
 * @author Cruz
 * @Date   17:19:00 30/2/2017
 * @version 01-00
 */
public class CellDefinition {
	/**
	 * 列号
	 */
<<<<<<< HEAD
	private int[] col;
=======
	private int col;
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	/**
	 * 行号
	 */
	private int[] row;
	/**
	 * 是否合并单元格
	 */
	private boolean merge;
	
	/**
	 * 列名-中文
	 */
    private String columnName;
    /**
     * 是否必须
     */
    private boolean required;
    /**
     * 单元格样式
     */
    private CellStyle cellStyle;
    /**
     * 字体样式
     */
    private FontStyle FontStyle;
    /**
     * 模板值之间是否有间隔
     */
    private boolean broken;
    /**
     * 属性
     */
    private Field field;
    /**
     * 属性注解
     */
    private EntityAttribute annotation;
<<<<<<< HEAD
    
	public int[] getCol() {
		return col;
	}
	public void setCol(int[] col) {
		this.col = col;
	}
	
=======
	public int getCol() {
		return col;
	}
	public void setCol(int col) {
		this.col = col;
	}
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	public int[] getRow() {
		return row;
	}
	public void setRow(int[] row) {
		this.row = row;
	}
	public boolean isMerge() {
		return merge;
	}
	public void setMerge(boolean merge) {
		this.merge = merge;
	}
	public String getColumnName() {
		return columnName;
	}
	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}
	public boolean isRequired() {
		return required;
	}
	public void setRequired(boolean required) {
		this.required = required;
	}
	public CellStyle getCellStyle() {
		return cellStyle;
	}
	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
	public FontStyle getFontStyle() {
		return FontStyle;
	}
	public void setFontStyle(FontStyle fontStyle) {
		FontStyle = fontStyle;
	}
	public boolean isBroken() {
		return broken;
	}
	public void setBroken(boolean broken) {
		this.broken = broken;
	}
	public Field getField() {
		return field;
	}
	public void setField(Field field) {
		this.field = field;
	}
	public EntityAttribute getAnnotation() {
		return annotation;
	}
	public void setAnnotation(EntityAttribute annotation) {
		this.annotation = annotation;
	}
    
    
}
