package com.ld.datacenter.poi.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ld.datacenter.dto.CollectionContrastDto;

/**
 * encapsulate common action for export flow
 * 
 * @author Cruz
 * @version 01-00
 */
public abstract class ExportBuilder {
	/**
	 * ultimately output stream
	 */
	private InputStream inputStream;
	/**
	 * XSSFWorkbook
	 */
	private XSSFWorkbook wb;

	private XSSFSheet sheet;
	/**
	 * export template name
	 */
	private String fileName;
	/**
	 * sheet name
	 */
	private String sheetName;

	public InputStream writeStreamFromTemplate(Class<?> cls)
			throws IOException {
		initWorkbook();
		setSheet();
<<<<<<< HEAD
		setTitle(cls);
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
		setCellValues(cls);
		extractPicture();
		outputStreamResult();
		closeStream();
		return this.inputStream;
	}

	// workbook
	abstract void initWorkbook() throws IOException;

	// sheetName
	abstract void setSheet();
<<<<<<< HEAD
	
	// title
	abstract void setTitle(Class<?> cls);
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf

	// selCellValue
	abstract void setCellValues(Class<?> cls);

	// setPicture
	abstract void extractPicture() throws IOException;

	// assign output
	protected void outputStreamResult() throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		wb.write(baos);
		this.inputStream = new ByteArrayInputStream(baos.toByteArray());
		baos.close();
	}

	// close stream
	protected void closeStream() throws IOException {
		this.inputStream.close();
	}

	/**
	 * 设置连续列excel内容
	 * 
	 * @param chartData
	 * @param beginCol
	 * @param endCol
	 * @param beginListIndex
	 * @param interval
	 * @return
	 */
	protected Map<Integer, String> buildExcelMap(
			List<CollectionContrastDto> chartData, int beginCol, int endCol,
			int beginListIndex, int interval) {
		Objects.requireNonNull(chartData, "input data must not bu null!");
		Map<Integer, String> kVal = new HashMap<Integer, String>();
		int i = 0;
		for (; beginCol <= endCol; beginCol++, i++) {
			kVal.putIfAbsent(beginCol,
					chartData.get(beginListIndex + interval * i).getCount()
							+ "");
		}
		return kVal;
	}

	/**
	 * 设置间隔列excel内容
	 * 
	 * @param chartData
	 * @param beginCol
	 * @param endCol
	 * @param beginListIndex
	 * @param interval
	 * @return
	 */
	protected Map<Integer, String> buildBrokenExcelMap(
			List<CollectionContrastDto> chartData, int[] cols,
			int beginListIndex, int interval) {
		Objects.requireNonNull(chartData, "input data must not bu null!");
		Map<Integer, String> kVal = new HashMap<Integer, String>();
		int i = 0, j = 0, max = cols.length;
		for (; i < max; i++, j++) {
			kVal.putIfAbsent(cols[i],
					chartData.get(beginListIndex + interval * j).getCount()
							+ "");
		}
		return kVal;
	}

	protected <T> String getCount(Class<T> cls, Object t) {
		if (cls == null || t == null) {
			throw new IllegalArgumentException(
					"the class type or object must not be null!");
		}
		try {
			Method m = cls.getDeclaredMethod("getCount");
			Object result=m.invoke(t, (Object[])null);
			return result.toString();
		} catch (Exception e) {
			e.printStackTrace();
			throw new IllegalArgumentException(e.getLocalizedMessage()
					+ ",the method must not be null!");
		}
	}

	public InputStream getInputStream() {
		return inputStream;
	}

	public void setInputStream(InputStream inputStream) {
		this.inputStream = inputStream;
	}

	public XSSFWorkbook getWb() {
		return wb;
	}

	public void setWb(XSSFWorkbook wb) {
		this.wb = wb;
	}

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public XSSFSheet getSheet() {
		return sheet;
	}

	public void setSheet(XSSFSheet sheet) {
		this.sheet = sheet;
	}

}
