package com.ld.datacenter.poi.facade;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.ArrayUtils;

import com.ld.datacenter.dto.search.BaseSearchDto;
import com.ld.datacenter.poi.service.ExcelExportService;
import com.ld.datacenter.poi.service.ExcelExportService.Builder;
import com.ld.datacenter.poi.service.ExportBuilder;

/**
 * 1.package POI export function
 * 2.reveal required parameters
 * @author Cruz
 * @version 01-00
 * @Date 2017/4/27
 */
public class PoiExportFacade {
	
	/**
	 * export excel without any parameter
	 * @param baseSearchDto
	 * @param excelPath
	 * @param exportName
	 * @param data
	 * @param dataCls
	 * @param templateCls
	 * @param pictureArea
	 * @return {@code Optional<InputStream>}
	 * @throws IOException
	 */
	public static <T>  Optional<InputStream> export(final BaseSearchDto baseSearchDto,final String excelPath,final String exportName,final List<T> data,final Class<T> dataCls,final Class<?> templateCls,final int[] pictureArea) throws IOException{
		return export(baseSearchDto,excelPath,exportName,data,dataCls,templateCls,pictureArea,false,false);
	}
	
	/**
	 * export excel with year
	 * @param baseSearchDto
	 * @param excelPath
	 * @param exportName
	 * @param data
	 * @param dataCls
	 * @param templateCls
	 * @param pictureArea
	 * @param hasYear
	 * @return {@code Optional<InputStream>}
	 * @throws IOException
	 */
	public static <T> Optional<InputStream> export(final BaseSearchDto baseSearchDto,final String excelPath,final String exportName,final List<T> data,final Class<T> dataCls,final Class<?> templateCls,final int[] pictureArea,boolean hasYear) throws IOException{
		return export(baseSearchDto,excelPath,exportName,data,dataCls,templateCls,pictureArea,true,false);
	}
	
	/**
	 * export excel with year and quarter
	 * @param baseSearchDto
	 * @param excelPath
	 * @param exportName
	 * @param data
	 * @param dataCls
	 * @param templateCls
	 * @param pictureArea
	 * @param hasYear
	 * @param hasQuarter
	 * @return {@code Optional<InputStream>}
	 * @throws IOException
	 */
	public static <T> Optional<InputStream> export(final BaseSearchDto baseSearchDto,final String excelPath,final String exportName,final List<T> data,final Class<T> dataCls,final Class<?> templateCls,final int[] pictureArea,boolean hasYear,boolean hasQuarter) throws IOException{
		Objects.requireNonNull(baseSearchDto, "baseSearchDto must not be null!");
		Builder<T> builder=new ExcelExportService.Builder<T>(excelPath,exportName,data,dataCls);
		String svg=Objects.requireNonNull(baseSearchDto.getSvgString(), "baseSearchDto svg string must not be null!");
		builder.setSvg(svg);
		List<String> _titles=baseSearchDto.getTitles();
		if(!CollectionUtils.isEmpty(_titles)){
			builder.setTitles(_titles);
		}
		if(hasYear){
			Integer _year=Objects.requireNonNull(baseSearchDto.getYear(), "baseSearchDto year must not be null!");
			builder.setYear(_year.toString());
		}
		if(hasQuarter){
			Integer _quarter=Objects.requireNonNull(baseSearchDto.getQuarter(), "baseSearchDto quarter must not be null!");
			builder.setQuarter(_quarter.toString());
		}
		org.springframework.util.Assert.isTrue(!ArrayUtils.isEmpty(pictureArea),"picture area have not been specified!");
		builder.setPictureStartCol(pictureArea[0]);
		builder.setPictureEndCol(pictureArea[1]);
		builder.setPictureStartRow(pictureArea[2]);
		builder.setPictureEndRow(pictureArea[3]);
		ExportBuilder exportBuilder=builder.build();
		return Optional.of(exportBuilder.writeStreamFromTemplate(templateCls));
	}
}
