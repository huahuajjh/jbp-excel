package com.excel.util.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 被贴上该注解的属性，将会被IExcelInput对象序列化
 * 
 * @author Bless
 * @version 1.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface InputColAnnotation {
	/**
	 * 表示该属性的数据在那行， 行坐标 从0开始 默认值为 -1 (-1表示该数据所在行为自适应)
	 */
	public int rowCoord() default -1;

	/**
	 * 表示该属性的数据在那列，列坐标 从0开始
	 */
	public int colCoord();

	/**
	 * 该数据自定义转换的方法名。（注意必须且只有一个参数是String类型）该方法不需要返回值。
	 *  public或者private都可以。 
	 *  方法结构：(public/private) String xxx(){ ...; }
	 */
	public String convertValMethod() default "";
	
	/**
	 * 是否必填字段，true为必填，false为不必填，默认为false
	 */
	public boolean required() default false;
	
	/**
	 * 是否唯一
	 */
	public boolean only() default false;
	
	/**
	 * 转换错误描述。
	 */
	public String converErrorMsg() default "转换错误";
	
	public String requiredErrorMsg() default "不可为空";
	
	public String onlyErrorMsg() default "必须唯一";
	
	/**
	 * 是否抛弃跨列数据
	 * @return
	 * 
	 * 默认为false
	 */
	public boolean isAbandonColspanData() default false;
	
	/**
	 * 是否抛弃跨行数据
	 * @return
	 * 
	 * 默认为false
	 */
	public boolean isAbandonRowspanData() default false;

}
