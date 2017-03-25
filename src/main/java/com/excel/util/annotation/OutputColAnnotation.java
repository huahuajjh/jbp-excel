package com.excel.util.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 被贴上该注解的属性，将会被IExcelOutput对象序列化
 * 
 * @author Bless
 * @version 1.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface OutputColAnnotation {
	/**
	 * 表示该属性的数据在那行， 行坐标 默认值为 -1 (-1表示该数据所在行为自适应)
	 */
	public int rowCoord() default -1;

	/**
	 * 表示该属性的数据在那列，列坐标 从0开始
	 */
	public int colCoord();

	/**
	 * 当前数据所占列单元格数量，默认在所在的列里 占据 1个单元格。(默认值为 1)
	 * 从1开始。
	 */
	public int colspan() default 1;

	/**
	 * 当前数据所占行单元格数量，默认在所在的行里 占据 1个单元格。(默认值为 1)
	 * 从1开始。
	 */
	public int rowspan() default 1;

	/**
	 * 数据的展示样式，填写样式的名称。默认为空。如果为空就使用系统默认样式
	 */
	public String styleDataName() default "";

	/**
	 * 自定义转换数据，传入方法名称（注意必须是无参的方法）。
	 * 方法 public或者private都可以。 
	 * 方法结构：
	 * (public/private) void xxx(String value){ ...; }
	 */
	public String convertValMethod() default "";
	
	/**
	 * 执行的函数
	 * @return
	 */
	public String formula() default "";
}
