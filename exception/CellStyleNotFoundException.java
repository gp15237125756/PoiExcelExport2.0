package com.ld.datacenter.poi.exception;
/**
 * override annotation exception
 * @author Cruz
 * @version 01-00
 * @Date 2017/3/31 10:10:00
 */
public class CellStyleNotFoundException extends RuntimeException{

	private static final long serialVersionUID = 1845298831506720372L;
	
	public CellStyleNotFoundException(){
		super();
	}
	
	
	public CellStyleNotFoundException(String message){
		super(message);
	}
	
	public CellStyleNotFoundException(Throwable Throwable){
		super(Throwable);
	}
	
	public CellStyleNotFoundException(String message,Throwable cause){
		super(message,cause);
	}
	
	public CellStyleNotFoundException(Class<?> cls,String message){
		super("entity ["+cls.getName()+"] "+message);
	}
}
