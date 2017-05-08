package com.ld.datacenter.poi.service;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.LinkedList;
import java.util.List;
import java.util.Objects;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import sun.misc.BASE64Decoder;

import com.ld.datacenter.poi.annotation.EntityAttribute;
import com.ld.datacenter.poi.annotation.ExcelEntity;
<<<<<<< HEAD
import com.ld.datacenter.poi.annotation.RenderDirection;
import com.ld.datacenter.poi.annotation.TitleAttribute;
import com.ld.datacenter.poi.annotation.TitleAttributes;
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
import com.ld.datacenter.poi.exception.AnnotationNotFoundException;
import com.ld.datacenter.poi.exception.IllegalTempClassException;
import com.ld.datacenter.poi.helper.CellDefinition;
import com.ld.datacenter.poi.helper.CellStyle;
import com.ld.datacenter.poi.helper.FontStyle;
<<<<<<< HEAD
import com.ld.datacenter.poi.helper.RenderStyle;
import com.ld.datacenter.poi.util.CellValueUtil;

/**
 * 注解方式实现导出功能
=======
import com.ld.datacenter.poi.util.CellValueUtil;

/**
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
 * @Description generate excel inputStream
 * @author Cruz
 * @version 01-00
 */
public class ExcelExportService extends ExportBuilder {
	// log
	private static final Log LOG = LogFactory.getLog(ExcelExportService.class);
	// year
	private final String year;
	// year
	private final String quarter;
	// month
	private final String month;
<<<<<<< HEAD
	//列数量
	private int COLUMN_NUMBER;
	//列数量 
	private int ROW_NUMBER;
	//标题数量
	private int TITLE_NUMBER;
	//数据渲染方向
	private RenderStyle RENDER_DIRECTION;
	
	private int ROWS;
	
	private int COLS;
=======
	
	private int COLUMN_NUMBER;
	
	private int ROWS;
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	//data
	private List<T> data;
	//data class  type
	private Class<T> dataType;
	// picture SVG format string
	private final String svgStr;
	// picture plot start column
	private final int pictureStartCol;
	// picture plot end column
	private final int pictureEndCol;
	// picture plot start row
	private final int pictureStartRow;
	// picture plot end row
	private final int pictureEndRow;
<<<<<<< HEAD
	//title values
	private final List<String> titles;
 
	@SuppressWarnings({ "unchecked", "rawtypes" })
=======

>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	public ExcelExportService(Builder builder) {
		// required
		this.setFileName(builder.fileName);
		// required
		this.setSheetName(builder.sheetName);
		// required
		this.setData(builder.data);
		//required
		this.setDataType(builder.dataType);
		
		/******************** optional parameter *******************/
		this.year=builder.year;
		this.quarter=builder.quarter;
		this.month=builder.month;
		this.svgStr = builder.svgStr;
		this.pictureStartCol = builder.pictureStartCol;
		this.pictureEndCol = builder.pictureEndCol;
		this.pictureStartRow = builder.pictureStartRow;
		this.pictureEndRow = builder.pictureEndRow;
<<<<<<< HEAD
		this.titles = builder.titles;
		
		/******************** optional parameter *******************/
	}

	@SuppressWarnings("hiding")
=======
		/******************** optional parameter *******************/
	}

>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
	public static class Builder<T> {
		private final String fileName;
		private final String sheetName;
		private final List<T> data;
		private Class<T> dataType;
		private String year;
		private String quarter;
		private String month;
		private String svgStr;
		private int pictureStartCol;
		private int pictureEndCol;
		private int pictureStartRow;
		private int pictureEndRow;
<<<<<<< HEAD
		//title values
		private List<String> titles;
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf

		public Builder(String fileName, String sheetName, List<T> data,Class<T> dataType) {
			this.fileName = fileName;
			this.sheetName = sheetName;
			this.data=data;
			this.dataType=dataType;
		}

		public Builder<T> setYear(String year) {
			this.year = year;
			return this;
		}

		public Builder<T> setQuarter(String quarter) {
			this.quarter = quarter;
			return this;
		}

		public Builder<T> setMonth(String month) {
			this.month = month;
			return this;
		}

		public Builder<T> setSvg(String svgStr) {
			this.svgStr = svgStr;
			return this;
		}

		public Builder<T> setPictureStartCol(int pictureStartCol) {
			this.pictureStartCol = pictureStartCol;
			return this;
		}

		public Builder<T> setPictureEndCol(int pictureEndCol) {
			this.pictureEndCol = pictureEndCol;
			return this;
		}

		public Builder<T> setPictureStartRow(int pictureStartRow) {
			this.pictureStartRow = pictureStartRow;
			return this;
		}

		public Builder<T> setPictureEndRow(int pictureEndRow) {
			this.pictureEndRow = pictureEndRow;
			return this;
		}
<<<<<<< HEAD
		
		public Builder<T> setTitles(List<String> titles) {
			this.titles = titles;
			return this;
		}
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf

		public ExcelExportService build() {
			return new ExcelExportService(this);
		}

	}

	protected void initWorkbook() throws IOException {
		String templatePath = Objects.requireNonNull(this.getFileName(),
				"template file name must not be null!");
		InputStream is=Objects.requireNonNull(this.getClass().getClassLoader().getResourceAsStream(templatePath),"create XSSFWorkbook failed,please check template's physical path!");
		this.setWb(new XSSFWorkbook(is));
	}

	@Override
	void setSheet() {
		XSSFWorkbook wb = Objects.requireNonNull(this.getWb(),
				"XSSFWorkbook must not be null!");
		XSSFSheet sheet = wb.getSheetAt(0);
		this.setSheet(sheet);
		this.setSheetName(this.getSheetName());
	}
<<<<<<< HEAD
	//set title according to template configure
	@Override
	void setTitle(Class<?> loadCls){
		if (loadCls == null) {
			if (LOG.isDebugEnabled()) {
				LOG.debug("loadCls must be Class type!");
			}
			throw new IllegalArgumentException(
					"loadCls must be Class type!");
=======

	// load template class ,then set cells
	@Override
	void setCellValues(Class<?> loadCls) {
		if (loadCls == null || !Class.class.isAssignableFrom(loadCls)) {
			if (LOG.isDebugEnabled()) {
				LOG.debug("loadCls must not be Class type!");
			}
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
		}
		ExcelEntity entity = loadCls.getAnnotation(ExcelEntity.class);
		if (entity == null) {
			if (LOG.isDebugEnabled()) {
<<<<<<< HEAD
				LOG.debug("[" + loadCls.getName() + "] must not be excel type!");
=======
				LOG.debug("[" + loadCls.getName() + "] must not be Class type!");
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
			}
			throw new AnnotationNotFoundException(loadCls,
					" is not a excel entity!");
		}
<<<<<<< HEAD
		TitleAttributes title = loadCls.getAnnotation(TitleAttributes.class);
		//set title if configured
		if (null!=title) {
			TitleAttribute[] titles=title.value();
			this.TITLE_NUMBER=titles.length;
			for(int i=0;i<TITLE_NUMBER;i++){
				TitleAttribute t=titles[i];
				int row=t.row(),col=t.col();
				CellStyle style=t.cellStyle();
				FontStyle font=t.fontStyle();
				CellValueUtil.setCell(this.getWb(),super.getSheet(),new int[]{row,col},style,font,this.titles.get(i));
			}
		}
	}
	
	// load template class ,then set cells
	@Override
	void setCellValues(Class<?> loadCls) {
		ExcelEntity entity = loadCls.getAnnotation(ExcelEntity.class);
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
		//tangle year,quarter,or month columns
		tangleHeader(entity);
		//load attributes type definition
		List<CellDefinition> contentDefinitionList=_getEntityFields(loadCls);
		if(CollectionUtils.isEmpty(contentDefinitionList)){
			throw new IllegalTempClassException(loadCls,
					" is not a legal entity!");
		}
<<<<<<< HEAD
		//get data render way which determine the iterate order of data set
 		RenderDirection direction=loadCls.getAnnotation(RenderDirection.class);
		if(null==direction){
			this.RENDER_DIRECTION=RenderStyle.HORIZATION;
		}else{
			this.RENDER_DIRECTION=direction.value();
		}
		//horizon
		if(this.RENDER_DIRECTION.equals(RenderStyle.HORIZATION)){
			this.COLUMN_NUMBER=contentDefinitionList.size();
			/****************** horizon direction render  ********************/
			for(int i=0;i<COLUMN_NUMBER;i++){
				//inner loop specify one type per time
				CellDefinition def=contentDefinitionList.get(i);
				int[] col=def.getCol();
				int[] rows=def.getRow();
				CellStyle style=def.getCellStyle();
				FontStyle font=def.getFontStyle();
				for(int _row=0;_row<this.ROWS;_row++){
					CellValueUtil.setCell(this.getWb(),super.getSheet(),new int[]{rows[_row],col[0]},style,font,this.getCount(this.dataType, this.data.get(i*ROWS+_row)));
				}
			}
			/****************** horizon direction render  ********************/
		}else{
		//vertical
			this.ROW_NUMBER=contentDefinitionList.size();
			/****************** vertical direction draw  ********************/
			for(int _row=0;_row<this.ROW_NUMBER;_row++){
				CellDefinition def=contentDefinitionList.get(_row);
				int[] col=def.getCol();
				int[] rows=def.getRow();
				CellStyle style=def.getCellStyle();
				FontStyle font=def.getFontStyle();
				for(int i=0;i<COLS;i++){
					CellValueUtil.setCell(this.getWb(),super.getSheet(),new int[]{rows[0],col[i]},style,font,this.getCount(this.dataType, this.data.get(_row*COLS+i)));
				}
			}
			/****************** vertical direction draw  ********************/
		}
=======
		this.COLUMN_NUMBER=contentDefinitionList.size();
		/****************** vertical direction draw  ********************/
		/*for(int _row=0;_row<this.ROWS;_row++){
			for(int i=0;i<COLUMN_NUMBER;i++){
				CellDefinition def=contentDefinitionList.get(i);
				int col=def.getCol();
				int[] rows=def.getRow();
				CellStyle style=def.getCellStyle();
				FontStyle font=def.getFontStyle();
				CellValueUtil.setCell(this.getWb(),super.getSheet(),new int[]{rows[_row],col},style,font,this.getCount(this.dataType, this.data.get(_row*COLUMN_NUMBER+i)));
			}
		}*/
		/****************** vertical direction draw  ********************/
		
		/****************** horizon direction draw  ********************/
		for(int i=0;i<COLUMN_NUMBER;i++){
			//inner loop specify one type per time
			CellDefinition def=contentDefinitionList.get(i);
			int col=def.getCol();
			int[] rows=def.getRow();
			CellStyle style=def.getCellStyle();
			FontStyle font=def.getFontStyle();
			for(int _row=0;_row<this.ROWS;_row++){
				CellValueUtil.setCell(this.getWb(),super.getSheet(),new int[]{rows[_row],col},style,font,this.getCount(this.dataType, this.data.get(i*ROWS+_row)));
			}
		}
		/****************** horizon direction draw  ********************/
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
		
	}
	
	private List<CellDefinition> _getEntityFields(Class<?> loadCls){
		List<CellDefinition> definition=new LinkedList<CellDefinition>();
		Field[] fields=loadCls.getDeclaredFields();
		for(Field f:fields){
			if(!f.isAnnotationPresent(EntityAttribute.class)){
				continue;
			}
			CellDefinition define=new CellDefinition();
			EntityAttribute attr=f.getAnnotation(EntityAttribute.class);
			define.setRow(attr.row());
			define.setCol(attr.col());
			define.setAnnotation(attr);
			define.setBroken(attr.broken());
			define.setCellStyle(attr.cellStyle());
			define.setField(f);
			define.setFontStyle(attr.fontStyle());
			define.setMerge(attr.merge());
			define.setRequired(attr.required());
			definition.add(define);
		}
		return definition;
	}
	
	private void tangleHeader(ExcelEntity entity){
		int[] yrs = entity.yrPos(), quars = entity.quaPos(), mons = entity
				.monPos();
		this.ROWS=entity.row();
<<<<<<< HEAD
		this.COLS=entity.col();
=======
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
		CellStyle style = entity.yrCellStyle();
		FontStyle font=entity.yrFontStyle();
		// set year value
		if (yrs!=null&&yrs.length > 1) {
			CellValueUtil.setCell(this.getWb(),super.getSheet(),yrs,style,font,this.year);
		}

		if (quars!=null&&quars.length > 1) {
			CellValueUtil.setCell(this.getWb(),super.getSheet(),quars,style,font,this.quarter);
		}

		if (mons!=null&&mons.length > 1) {
			CellValueUtil.setCell(this.getWb(),super.getSheet(),mons,style,font,this.month);
		}
	}

	@Override
	void extractPicture() throws IOException {
		this.extractPicturePortion(svgStr, this.getWb(), this.getSheet(),
				pictureStartCol, pictureEndCol, pictureStartRow, pictureEndRow);
	}

	/**
	 * 抽象出图片生成业务代码
	 * 
	 * @throws IOException
	 */
	private void extractPicturePortion(String svgString, XSSFWorkbook wb,
			XSSFSheet sheet, int startCol, int endCol, int startRow, int endRow)
			throws IOException {
		// 图片
		if (org.apache.commons.lang3.StringUtils.isNotBlank(svgString)) {
			byte[] safeDataBytes = new BASE64Decoder().decodeBuffer(svgString);
			int pictureIdx = wb.addPicture(safeDataBytes,
					Workbook.PICTURE_TYPE_JPEG);
			CreationHelper helper = wb.getCreationHelper();
			// Create the drawing patriarch. This is the top level container for
			// all shapes.
			Drawing drawing = sheet.createDrawingPatriarch();
			// add a picture shape
			ClientAnchor anchor = helper.createClientAnchor();
			// set top-left corner of the picture,
			// subsequent call of Picture#resize() will operate relative to it
			anchor.setCol1(startCol);
			anchor.setCol2(endCol);
			anchor.setRow1(startRow);
			anchor.setRow2(endRow);
			anchor.setDx1(0);
			anchor.setDy1(0);
			anchor.setDx2(0);
			anchor.setDy2(0);
			anchor.setAnchorType(ClientAnchor.MOVE_DONT_RESIZE);
			Picture pict = drawing.createPicture(anchor, pictureIdx);
			pict.resize(1);
		}
	}

	public List<T> getData() {
		return data;
	}

	public void setData(List<T> data){
		this.data=data;
	}

	public Class<T> getDataType() {
		return dataType;
	}

	public void setDataType(Class<T> dataType) {
		this.dataType = dataType;
	}
	
	
}
