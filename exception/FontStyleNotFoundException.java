package com.ld.datacenter.poi.exception;
/**
 * override annotation exception
 * @author Cruz
 * @version 01-00
 * @Date 2017/3/31 10:10:00
 */
public class FontStyleNotFoundException extends RuntimeException{

	private static final long serialVersionUID = 1845298831506720372L;
	
	public FontStyleNotFoundException(){
		super();
	}
	
	
	public FontStyleNotFoundException(String message){
		super(message);
	}
	
	public FontStyleNotFoundException(Throwable Throwable){
		super(Throwable);
	}
	
	public FontStyleNotFoundException(String message,Throwable cause){
		super(message,cause);
	}
	
	public FontStyleNotFoundException(Class<?> cls,String message){
		super("entity ["+cls.getName()+"] "+message);
	}
}
