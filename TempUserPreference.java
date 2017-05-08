package com.ld.datacenter.template.guide;

import com.ld.datacenter.poi.annotation.EntityAttribute;
import com.ld.datacenter.poi.annotation.ExcelEntity;
import com.ld.datacenter.poi.annotation.RenderDirection;
import com.ld.datacenter.poi.helper.CellStyle;
import com.ld.datacenter.poi.helper.FontStyle;
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
	
/*	@EntityAttribute(columnName="其他",required=true,col=14,row={3},merge=false,broken=false,cellStyle=CellStyle.STANDARD,fontStyle=FontStyle.COLOR)
	private String  others;*/
	
	
}
