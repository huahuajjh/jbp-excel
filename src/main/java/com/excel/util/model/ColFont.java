package com.excel.util.model;

/**
 * 单元格的字体样式
 * 
 * @author Bless
 * @version 1.0
 */
public class ColFont {
	/**
	 * Normal boldness (not bold)
	 */

	public final static short BOLDWEIGHT_NORMAL = 0x190;

	/**
	 * Bold boldness (bold)
	 */

	public final static short BOLDWEIGHT_BOLD = 0x2bc;

	/**
	 * normal type of black color.
	 */

	public final static short COLOR_NORMAL = 0x7fff;

	/**
	 * Dark Red color
	 */

	public final static short COLOR_RED = 0xa;

	/**
	 * no type offsetting (not super or subscript)
	 */

	public final static short SS_NONE = 0;

	/**
	 * superscript
	 */

	public final static short SS_SUPER = 1;

	/**
	 * subscript
	 */

	public final static short SS_SUB = 2;

	/**
	 * not underlined
	 */

	public final static byte U_NONE = 0;

	/**
	 * single (normal) underline
	 */

	public final static byte U_SINGLE = 1;

	/**
	 * double underlined
	 */

	public final static byte U_DOUBLE = 2;

	/**
	 * accounting style single underline
	 */

	public final static byte U_SINGLE_ACCOUNTING = 0x21;

	/**
	 * accounting style double underline
	 */

	public final static byte U_DOUBLE_ACCOUNTING = 0x22;

	/**
	 * ANSI character set
	 */
	public final static byte ANSI_CHARSET = 0;

	/**
	 * Default character set.
	 */
	public final static byte DEFAULT_CHARSET = 1;

	/**
	 * Symbol character set
	 */
	public final static byte SYMBOL_CHARSET = 2;

	/**
	 * set the name for the font (i.e. Arial)
	 * 
	 * @param name
	 *            String representing the name of the font to use
	 */

	public void setFontName(String name) {
		this.fontName = name;
	}

	private String fontName = "Arial";

	/**
	 * get the name for the font (i.e. Arial)
	 * 
	 * @return String representing the name of the font to use
	 */
	public String getFontName() {
		return this.fontName;
	}

	/**
	 * set the font height in unit's of 1/20th of a point. Maybe you might want
	 * to use the setFontHeightInPoints which matches to the familiar 10, 12, 14
	 * etc..
	 * 
	 * @param height
	 *            height in 1/20ths of a point
	 * @see #setFontHeightInPoints(short)
	 */

	public void setFontHeight(short height) {
		this.fontHeight = height;
	}

	private short fontHeight = 0xc8;

	/**
	 * get the font height in unit's of 1/20th of a point. Maybe you might want
	 * to use the getFontHeightInPoints which matches to the familiar 10, 12, 14
	 * etc..
	 * 
	 * @return short - height in 1/20ths of a point
	 * @see #getFontHeightInPoints()
	 */
	public short getFontHeight() {
		return this.fontHeight;
	}

	/**
	 * set the font height
	 * 
	 * @param height
	 *            height in the familiar unit of measure - points
	 * @see #setFontHeight(short)
	 */

	public void setFontHeightInPoints(short height) {
		this.fontHeightInPoints = height;
	}

	private short fontHeightInPoints = 11;

	/**
	 * get the font height
	 * 
	 * @return short - height in the familiar unit of measure - points
	 * @see #getFontHeight()
	 */
	public short getFontHeightInPoints() {
		return fontHeightInPoints;
	}

	/**
	 * set whether to use italics or not
	 * 
	 * @param italic
	 *            italics or not
	 */
	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	private boolean italic = false;

	/**
	 * get whether to use italics or not
	 * 
	 * @return italics or not
	 */
	public boolean getItalic() {
		return this.italic;
	}

	/**
	 * set whether to use a strikeout horizontal line through the text or not
	 * 
	 * @param strikeout
	 *            or not
	 */

	public void setStrikeout(boolean strikeout) {
		this.strikeout = strikeout;
	}

	private boolean strikeout = false;

	/**
	 * get whether to use a strikeout horizontal line through the text or not
	 * 
	 * @return strikeout or not
	 */

	public boolean getStrikeout() {
		return this.strikeout;
	}

	/**
	 * set the color for the font
	 * 
	 * @param color
	 *            to use
	 * @see #COLOR_NORMAL Note: Use this rather than HSSFColor.AUTOMATIC for
	 *      default font color
	 * @see #COLOR_RED
	 */

	public void setColor(short color) {
		this.color = color;
	}

	private short color = COLOR_NORMAL;

	/**
	 * get the color for the font
	 * 
	 * @return color to use
	 * @see #COLOR_NORMAL
	 * @see #COLOR_RED
	 * @see org.apache.poi.hssf.usermodel.HSSFPalette#getColor(short)
	 */
	public short getColor() {
		return this.color;
	}

	/**
	 * set normal,super or subscript.
	 * 
	 * @param offset
	 *            type to use (none,super,sub)
	 * @see #SS_NONE
	 * @see #SS_SUPER
	 * @see #SS_SUB
	 */

	public void setTypeOffset(short offset) {
		this.offset = offset;
	}

	private short offset = SS_NONE;

	/**
	 * get normal,super or subscript.
	 * 
	 * @return offset type to use (none,super,sub)
	 * @see #SS_NONE
	 * @see #SS_SUPER
	 * @see #SS_SUB
	 */

	public short getTypeOffset() {
		return this.offset;
	}

	/**
	 * set type of text underlining to use
	 * 
	 * @param underline
	 *            type
	 * @see #U_NONE
	 * @see #U_SINGLE
	 * @see #U_DOUBLE
	 * @see #U_SINGLE_ACCOUNTING
	 * @see #U_DOUBLE_ACCOUNTING
	 */

	public void setUnderline(byte underline) {
		this.underline = underline;
	}

	private byte underline = U_NONE;

	/**
	 * get type of text underlining to use
	 * 
	 * @return underlining type
	 * @see #U_NONE
	 * @see #U_SINGLE
	 * @see #U_DOUBLE
	 * @see #U_SINGLE_ACCOUNTING
	 * @see #U_DOUBLE_ACCOUNTING
	 */

	public byte getUnderline() {
		return this.underline;
	}

	/**
	 * get character-set to use.
	 * 
	 * @return character-set
	 * @see #ANSI_CHARSET
	 * @see #DEFAULT_CHARSET
	 * @see #SYMBOL_CHARSET
	 */
	public int getCharSet() {
		return this.charSet;
	}

	private Byte charSet = ANSI_CHARSET;

	/**
	 * set character-set to use.
	 * 
	 * @see #ANSI_CHARSET
	 * @see #DEFAULT_CHARSET
	 * @see #SYMBOL_CHARSET
	 */
	public void setCharSet(byte charSet) {
		this.charSet = charSet;
	}

	/**
	 * set character-set to use.
	 * 
	 * @see #ANSI_CHARSET
	 * @see #DEFAULT_CHARSET
	 * @see #SYMBOL_CHARSET
	 */
	public void setCharSet(int charSet) {
		this.charSet = (byte)charSet;
	}

	public void setBoldweight(short boldweight) {
		this.boldweight = boldweight;
	}

	private short boldweight = 0x190;

	public short getBoldweight() {
		return boldweight;
	}

	public void setBold(Boolean bold) {
		this.bold = bold;
	}

	private Boolean bold = false;

	public Boolean getBold() {
		return bold;
	}
}
