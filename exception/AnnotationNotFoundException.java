package com.ld.datacenter.poi.exception;
/**
 * override annotation exception
 * @author Cruz
 * @version 01-00
 * @Date 2017/3/31 10:10:00
 */
public class AnnotationNotFoundException extends RuntimeException{

	private static final long serialVersionUID = 1845298831506720372L;
	
	public AnnotationNotFoundException(){
		super();
	}
	
	
	public AnnotationNotFoundException(String message){
		super(message);
	}
	
	public AnnotationNotFoundException(Throwable Throwable){
		super(Throwable);
	}
	
	public AnnotationNotFoundException(String message,Throwable cause){
		super(message,cause);
	}
	
	public AnnotationNotFoundException(Class<?> cls,String message){
		super("entity ["+cls.getName()+"] "+message);
	}
}
