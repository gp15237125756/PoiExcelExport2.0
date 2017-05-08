# PoiExcelExport


<<<<<<< HEAD
PoiExcelExport实现了[POI](http://poi.apache.org/)对`.xlsx`文件的導出功能的封装，实现了根據导出excel模板类定义，自动填充excel的功能。

该类库可以增加自定义cell样式和字体样式进行扩展。

# Notion

- 该poi库使用注解类作为输出excel模板，如有需要，可以使用xml格式的模板。
- 该类库适用于导出模板固定情况下，如普通的分类。
- 在导出类别动态变化的情况下，如常展厅、临展厅这类不固定分类场景下，不要使用。


## Scenario
=======
PoiExcelExport实现了`Java POI`对`xlsx`文件的導出功能的封装，实现了根據导出excel模板类定义，自动填充excel的功能。

该类库可以增加自定义cell样式和字体样式进行扩展。

# 注意

该poi库使用注解类作为输出excel模板，如有需要，可以使用xml格式的模板。


## 应用场景
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf

  该类库主要用应用场景是在数据中心做导出功能，目前业务较为简单，所以没有对导出到excel模板显示的数据做类型校验，后期可以进行扩展。


## 依赖

<<<<<<< HEAD
PoiExcelExport依赖于**Apache POI 3.8**类库。

## 使用说明

参考 *com.ld.datacenter.service/CollectionStatisticService/exportMonthly()(line 1398)* 测试用例。

  
***
	 List<GuideUserPreferenceResultDto> chartData=this.searchGuideUserPreference(baseSearchDto);
	//export all user preference
	PoiExportFacade.export(baseSearchDto, DataCenterConstant.EXPORT_ALL_USER_PREFERENCE_TEMP_PATH, "全部用户喜好统计", chartData, GuideUserPreferenceResultDto.class, TempUserPreference.class, new int[]{1,14,5,17});
			
=======
PoiExcelExport依赖于`Apache POI 3.8`类库。

## 使用说明

- 使用示例请参考`com.ld.datacenter.service/CollectionStatisticService/exportCategory()`测试用例。

  
***
  
	List<CollectionContrastDto> chartData=this.collectionCategoryQuarterContrast(year);
			ExportBuilder builder=new ExcelExportService.Builder<CollectionContrastDto>("/com/ld/datacenter/service/馆藏分类季对比.xlsx","馆藏分类季度对比",chartData,CollectionContrastDto.class)
					.setYear(year)
					.setSvg(svgCode)
					.setPictureStartCol(1)
					.setPictureEndCol(14)
					.setPictureStartRow(8)
					.setPictureEndRow(20)
					.build();
			return Optional.of(builder.writeStreamFromTemplate(TempCategory.class));


>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
***

### 使用方式

<<<<<<< HEAD
- 1.*取得数据*

       	List<GuideUserPreferenceResultDto> chartData=this.searchGuideUserPreference(baseSearchDto);
       
- 2.*输入参数，构建输出流*

			PoiExportFacade.export(baseSearchDto, DataCenterConstant.EXPORT_ALL_USER_PREFERENCE_TEMP_PATH, "全部用户喜好统计", chartData, GuideUserPreferenceResultDto.class, TempUserPreference.class, new int[]{1,14,5,17});

### 注册新的样式

* 注册的新的cell样式类型类必须 实现 [ICellStyleFactory](位于com.ld.datacenter.poi.font)接口。

* 注册的新的cell字体类型类必须实现  [IFontFactory](位于com.ld.datacenter.poi.style)接口。

* 添加cell格子样式,修改[XSSFCellStyleLib](位于com.ld.datacenter.poi.extention)类，增加以下代码 
=======
- 1.取得数据

       List<CollectionContrastDto> chartData=this.collectionCategoryQuarterContrast(year);
       
- 2.输入参数，构建输出流

			ExportBuilder builder=new ExcelExportService.Builder<CollectionContrastDto>("/com/ld/datacenter/service/馆藏分类季度对比.xlsx","馆藏分类季度对比",chartData,CollectionContrastDto.class)
					.setYear(year)
					.setSvg(svgCode)
					.setPictureStartCol(1)
					.setPictureEndCol(14)
					.setPictureStartRow(8)
					.setPictureEndRow(20)
					.build();
			return Optional.of(builder.writeStreamFromTemplate(TempCategory.class));


### 注册新的样式

- 注册的新的cell样式类型类必须 实现  `ICellStyleFactory` 接口。

- 注册的新的cell字体类型类必须实现  `IFontFactory` 接口。

- 添加cell格子样式,修改XSSFCellStyleLib类，增加以下代码 
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
  
  	
     	convert.convert(CellStyle.STANDARD, new TextCellStyleImpl(wb),this);


	
<<<<<<< HEAD
例如，[XSSFFontStyleLib](位于com.ld.datacenter.poi.extention)中添加以下代码，表示增加了三种格子样式：
=======
例如，XSSFFontStyleLib中添加以下代码，表示增加了三种格子样式：
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf

		 
			 ICellStyleConvert convert=new CellStyleConvert(); <br />  
			 //添加cell樣式 <br />  
			 convert.convert(CellStyle.NONE, new DefaultCellStyleImpl(wb));<br />  
			 convert.convert(CellStyle.DEFAULT, new DefaultCellStyleImpl(wb));<br />  
			 convert.convert(CellStyle.STANDARD, new TextCellStyleImpl(wb));<br />  
		 
<<<<<<< HEAD
增加新的样式或字体，分别在以下类中扩展。

- [XSSFCellUtil](com.ld.datacenter.poi.util.XSSFCellUtil)
- [XSSFontUtil](com.ld.datacenter.poi.util.XSSFontUtil)
		

### 实体对象

实体类必须标注**@ExcelEntity**注解，其中row定义需要填充总行数，col对应需要填充总列数


 同时需要填充的字段标注**@EntityAttribute**注解



	/**
	 * 用户喜好模板
	 * @author Cruz
	 * @Date 2017/4/27
	 * @version 01-00
	 */
	import com.ld.datacenter.poi.helper.RenderStyle;
	@ExcelEntity(row=1)
	@RenderDirection(RenderStyle.HORIZATION)
	public class TempUserPreference {
	@EntityAttribute(columnName="70后艺术",required=true,col=2,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String afterSeventy;
	
	@EntityAttribute(columnName="瓷杂",required=true,col=3,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  porcelainhethrd;
	
	@EntityAttribute(columnName="古代书画",required=true,col=4,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  ancientPatings;
	
	@EntityAttribute(columnName="红色经典",required=true,col=5,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  redClassic;
	
	@EntityAttribute(columnName="近现代书画",required=true,col=6,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  modernPainting;
	
	@EntityAttribute(columnName="老油画",required=true,col=7,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  oldOilPainting;
	
	@EntityAttribute(columnName="连环画",required=true,col=8,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  comicsandother;
	
	@EntityAttribute(columnName="外国艺术",required=true,col=9,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  otherForeignArt;
	
	@EntityAttribute(columnName="现当代书画",required=true,col=10,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  contemporaryArt;
	
	@EntityAttribute(columnName="亚洲艺术",required=true,col=11,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  asianArt;
	
	@EntityAttribute(columnName="国外古董",required=true,col=12,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  foreignAntique;
	
	@EntityAttribute(columnName="当代水墨",required=true,col=13,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  contemporaryInk;
	
	
	
}
	
	
}

# 2017-4-27 update
- 增加外观，屏蔽导出细节，进一步简化导出
				
				
				public Optional<InputStream> exportUserPreference(final BaseSearchDto baseSearchDto) throws IOException{
						if(BasicConstants.COLLECTION_CONTRAST_OPTION_ALL.equals(baseSearchDto.getContrastProp())){
							List<GuideUserPreferenceResultDto> chartData=this.searchGuideUserPreference(baseSearchDto);
							//export all user preference
							if(null==baseSearchDto.getYear()){
								return PoiExportFacade.export(baseSearchDto, DataCenterConstant.EXPORT_ALL_USER_PREFERENCE_TEMP_PATH, "全部用户喜好统计", chartData, GuideUserPreferenceResultDto.class, TempUserPreference.class, new int[]{1,14,5,17});
							}else{
								//export specified year's user preference
								if(baseSearchDto.getQuarter()!=null)
									return PoiExportFacade.export(baseSearchDto, DataCenterConstant.EXPORT_SPECIFIED_USER_PREFERENCE_TEMP_PATH, "用户喜好统计", chartData, GuideUserPreferenceResultDto.class, TempSpecifiedYearUserPreference.class, new int[]{1,14,5,17},true,true);
								else
									return PoiExportFacade.export(baseSearchDto, DataCenterConstant.EXPORT_SPECIFIED_USER_PREFERENCE_TEMP_PATH, "用户喜好统计", chartData, GuideUserPreferenceResultDto.class, TempSpecifiedYearUserPreference.class, new int[]{1,14,5,17},true);
							}
							
						}else if(BasicConstants.COLLECTION_CONTRAST_OPTION_QUARTER.equals(baseSearchDto.getContrastProp())){
							List<GuideUserPreferenceContrastDto> chartData=this.searchGuideUserPreferenceQuarterContr(baseSearchDto);
							return PoiExportFacade.export(baseSearchDto, DataCenterConstant.EXPORT_USER_PREFERENCE_QUARTER_CONTRAST_TEMP_PATH, "用户喜好季度对比", chartData, GuideUserPreferenceContrastDto.class, TempUserPreferenceQuarterContrast.class, new int[]{1,14,8,20},true);
						}else if(BasicConstants.COLLECTION_CONTRAST_OPTION_MONTH.equals(baseSearchDto.getContrastProp())){
							List<GuideUserPreferenceContrastDto> chartData=this.searchGuideUserPreferenceMonthContr(baseSearchDto);
							//12 months
							if(null==baseSearchDto.getQuarter()){
								return PoiExportFacade.export(baseSearchDto, DataCenterConstant.EXPORT_USER_PREFERENCE_MONTH_CONTRAST_TEMP_PATH, "用户喜好月份对比", chartData, GuideUserPreferenceContrastDto.class, TempUserPreferenceMonthContrast.class, new int[]{1,14,16,28},true);
							}else{
							//3 months
								return PoiExportFacade.export(baseSearchDto, DataCenterConstant.EXPORT_USER_PREFERENCE_QUARTER_MONTH_CONTRAST_TEMP_PATH, "用户喜好季度月份对比", chartData, GuideUserPreferenceContrastDto.class, TempUserPreferenceQuarterMonthContrast.class, new int[]{1,14,7,19},true,true);
							}
						}
						return null;
					}

- 增加导出标题功能，此参数封装在**BaseSearchDto**中，同时需要在模板**Tem**打头文件中，指定标题位置

				@TitleAttributes({
					@TitleAttribute(row=2,col=5,cellStyle=CellStyle.TITLE,fontStyle=FontStyle.TITLE),
					@TitleAttribute(row=2,col=7,cellStyle=CellStyle.TITLE,fontStyle=FontStyle.TITLE),
					@TitleAttribute(row=2,col=9,cellStyle=CellStyle.TITLE,fontStyle=FontStyle.TITLE)})





				
=======



### 实体对象

实体类必须标注@ExcelEntity注解， 同时需要填充的字段标注@EntityAttribute注解


	@ExcelEntity(row=4,yrPos={1,13},yrCellStyle=CellStyle.STANDARD)
	public class TempCategory {

	@EntityAttribute(columnName="70后艺术",required=true,col=2,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String afterSeventy;
	
	@EntityAttribute(columnName="瓷杂",required=true,col=3,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  porcelainhethrd;
	
	@EntityAttribute(columnName="古代书画",required=true,col=4,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  ancientPatings;
	
	@EntityAttribute(columnName="红色经典",required=true,col=5,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  redClassic;
	
	@EntityAttribute(columnName="近现代书画",required=true,col=6,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  modernPainting;
	
	@EntityAttribute(columnName="老油画",required=true,col=7,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  oldOilPainting;
	
	@EntityAttribute(columnName="连环画",required=true,col=8,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  comicsandother;
	
	@EntityAttribute(columnName="外国艺术",required=true,col=9,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  otherForeignArt;
	
	@EntityAttribute(columnName="现当代书画",required=true,col=10,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  contemporaryArt;
	
	@EntityAttribute(columnName="亚洲艺术",required=true,col=11,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  asianArt;
	
	@EntityAttribute(columnName="国外古董",required=true,col=12,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  foreignAntique;
	
	@EntityAttribute(columnName="当代水墨",required=true,col=13,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  contemporaryInk;
	
	@EntityAttribute(columnName="其他",required=true,col=14,row={3,4,5,6},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  others;
	
	
}
>>>>>>> fe2012a7f8558d8df36b789847bdc41c788d6eaf
