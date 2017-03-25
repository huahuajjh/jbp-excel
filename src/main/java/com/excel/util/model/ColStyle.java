package com.excel.util.model;

/**
 * 单元格的样式
 * 
 * @author Bless
 * @version 1.0
 */
public class ColStyle {
	/**
	 * general (normal) horizontal alignment
	 */
	public final static short ALIGN_GENERAL = 0x0;

	/**
	 * left-justified horizontal alignment
	 */
	public final static short ALIGN_LEFT = 0x1;

	/**
	 * center horizontal alignment
	 */
	public final static short ALIGN_CENTER = 0x2;

	/**
	 * right-justified horizontal alignment
	 */
	public final static short ALIGN_RIGHT = 0x3;

	/**
	 * fill? horizontal alignment
	 */
	public final static short ALIGN_FILL = 0x4;

	/**
	 * justified horizontal alignment
	 */
	public final static short ALIGN_JUSTIFY = 0x5;

	/**
	 * center-selection? horizontal alignment
	 */
	public final static short ALIGN_CENTER_SELECTION = 0x6;

	/**
	 * top-aligned vertical alignment
	 */
	public final static short VERTICAL_TOP = 0x0;

	/**
	 * center-aligned vertical alignment
	 */
	public final static short VERTICAL_CENTER = 0x1;

	/**
	 * bottom-aligned vertical alignment
	 */
	public final static short VERTICAL_BOTTOM = 0x2;

	/**
	 * vertically justified vertical alignment
	 */
	public final static short VERTICAL_JUSTIFY = 0x3;

	/**
	 * No border
	 */
	public final static short BORDER_NONE = 0x0;

	/**
	 * Thin border
	 */
	public final static short BORDER_THIN = 0x1;

	/**
	 * Medium border
	 */
	public final static short BORDER_MEDIUM = 0x2;

	/**
	 * dash border
	 */
	public final static short BORDER_DASHED = 0x3;

	/**
	 * dot border
	 */
	public final static short BORDER_HAIR = 0x7;

	/**
	 * Thick border
	 */
	public final static short BORDER_THICK = 0x5;

	/**
	 * double-line border
	 */
	public final static short BORDER_DOUBLE = 0x6;

	/**
	 * hair-line border
	 */
	public final static short BORDER_DOTTED = 0x4;

	/**
	 * Medium dashed border
	 */
	public final static short BORDER_MEDIUM_DASHED = 0x8;

	/**
	 * dash-dot border
	 */
	public final static short BORDER_DASH_DOT = 0x9;

	/**
	 * medium dash-dot border
	 */
	public final static short BORDER_MEDIUM_DASH_DOT = 0xA;

	/**
	 * dash-dot-dot border
	 */
	public final static short BORDER_DASH_DOT_DOT = 0xB;

	/**
	 * medium dash-dot-dot border
	 */
	public final static short BORDER_MEDIUM_DASH_DOT_DOT = 0xC;

	/**
	 * slanted dash-dot border
	 */
	public final static short BORDER_SLANTED_DASH_DOT = 0xD;

	/** No background */
	public final static short NO_FILL = 0;

	/** Solidly filled */
	public final static short SOLID_FOREGROUND = 1;

	/** Small fine dots */
	public final static short FINE_DOTS = 2;

	/** Wide dots */
	public final static short ALT_BARS = 3;

	/** Sparse dots */
	public final static short SPARSE_DOTS = 4;

	/** Thick horizontal bands */
	public final static short THICK_HORZ_BANDS = 5;

	/** Thick vertical bands */
	public final static short THICK_VERT_BANDS = 6;

	/** Thick backward facing diagonals */
	public final static short THICK_BACKWARD_DIAG = 7;

	/** Thick forward facing diagonals */
	public final static short THICK_FORWARD_DIAG = 8;

	/** Large spots */
	public final static short BIG_SPOTS = 9;

	/** Brick-like layout */
	public final static short BRICKS = 10;

	/** Thin horizontal bands */
	public final static short THIN_HORZ_BANDS = 11;

	/** Thin vertical bands */
	public final static short THIN_VERT_BANDS = 12;

	/** Thin backward diagonal */
	public final static short THIN_BACKWARD_DIAG = 13;

	/** Thin forward diagonal */
	public final static short THIN_FORWARD_DIAG = 14;

	/** Squares */
	public final static short SQUARES = 15;

	/** Diamonds */
	public final static short DIAMONDS = 16;

	/** Less Dots */
	public final static short LESS_DOTS = 17;

	/** Least Dots */
	public final static short LEAST_DOTS = 18;

	// -----------------------------------------------------------------------------------

	/**
	 * set the data format (must be a valid format)
	 */
	public void setDataFormat(String fmt) {
		this.fmt = fmt;
	}

	private String fmt = null;

	/**
	 * get the index of the format
	 */
	public String getDataFormat() {
		return this.fmt;
	}

	/**
	 * set the cell's using this style to be hidden
	 * 
	 * @param hidden
	 *            - whether the cell using this style should be hidden
	 */
	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	private boolean hidden = false;

	/**
	 * get whether the cell's using this style are to be hidden
	 * 
	 * @return hidden - whether the cell using this style should be hidden
	 */
	public boolean getHidden() {
		return this.hidden;
	}

	/**
	 * set the cell's using this style to be locked
	 * 
	 * @param locked
	 *            - whether the cell using this style should be locked
	 */
	public void setLocked(boolean locked) {
		this.locked = locked;
	}

	private boolean locked = false;

	/**
	 * get whether the cell's using this style are to be locked
	 * 
	 * @return hidden - whether the cell using this style should be locked
	 */
	public boolean getLocked() {
		return this.locked;
	}

	/**
	 * set the type of horizontal alignment for the cell
	 * 
	 * @param align
	 *            - the type of alignment
	 * @see #ALIGN_GENERAL
	 * @see #ALIGN_LEFT
	 * @see #ALIGN_CENTER
	 * @see #ALIGN_RIGHT
	 * @see #ALIGN_FILL
	 * @see #ALIGN_JUSTIFY
	 * @see #ALIGN_CENTER_SELECTION
	 */
	public void setAlignment(short align) {
		this.align = align;
	}

	private short align = ALIGN_GENERAL;

	/**
	 * get the type of horizontal alignment for the cell
	 * 
	 * @return align - the type of alignment
	 * @see #ALIGN_GENERAL
	 * @see #ALIGN_LEFT
	 * @see #ALIGN_CENTER
	 * @see #ALIGN_RIGHT
	 * @see #ALIGN_FILL
	 * @see #ALIGN_JUSTIFY
	 * @see #ALIGN_CENTER_SELECTION
	 */
	public short getAlignment() {
		return this.align;
	}

	/**
	 * Set whether the text should be wrapped. Setting this flag to
	 * <code>true</code> make all content visible whithin a cell by displaying
	 * it on multiple lines
	 *
	 * @param wrapped
	 *            wrap text or not
	 */
	public void setWrapText(boolean wrapped) {
		this.wrapped = wrapped;
	}

	private boolean wrapped = false;

	/**
	 * get whether the text should be wrapped
	 * 
	 * @return wrap text or not
	 */
	public boolean getWrapText() {
		return this.wrapped;
	}

	/**
	 * set the type of vertical alignment for the cell
	 * 
	 * @param align
	 *            the type of alignment
	 * @see #VERTICAL_TOP
	 * @see #VERTICAL_CENTER
	 * @see #VERTICAL_BOTTOM
	 * @see #VERTICAL_JUSTIFY
	 */
	public void setVerticalAlignment(short align) {
		this.verticalAlign = align;
	}

	private short verticalAlign = VERTICAL_CENTER;

	/**
	 * get the type of vertical alignment for the cell
	 * 
	 * @return align the type of alignment
	 * @see #VERTICAL_TOP
	 * @see #VERTICAL_CENTER
	 * @see #VERTICAL_BOTTOM
	 * @see #VERTICAL_JUSTIFY
	 */
	public short getVerticalAlignment() {
		return this.verticalAlign;
	}

	/**
	 * set the degree of rotation for the text in the cell
	 * 
	 * @param rotation
	 *            degrees (between -90 and 90 degrees)
	 */
	public void setRotation(short rotation) {
		this.rotation = rotation;
	}

	private short rotation = 0;

	/**
	 * get the degree of rotation for the text in the cell
	 * 
	 * @return rotation degrees (between -90 and 90 degrees)
	 */
	public short getRotation() {
		return this.rotation;
	}

	/**
	 * set the number of spaces to indent the text in the cell
	 * 
	 * @param indent
	 *            - number of spaces
	 */
	public void setIndention(short indent) {
		this.indent = indent;
	}

	private short indent = 0;

	/**
	 * get the number of spaces to indent the text in the cell
	 * 
	 * @return indent - number of spaces
	 */
	public short getIndention() {
		return this.indent;
	}

	/**
	 * set the type of border to use for the left border of the cell
	 * 
	 * @param border
	 *            type
	 * @see #BORDER_NONE
	 * @see #BORDER_THIN
	 * @see #BORDER_MEDIUM
	 * @see #BORDER_DASHED
	 * @see #BORDER_DOTTED
	 * @see #BORDER_THICK
	 * @see #BORDER_DOUBLE
	 * @see #BORDER_HAIR
	 * @see #BORDER_MEDIUM_DASHED
	 * @see #BORDER_DASH_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT
	 * @see #BORDER_DASH_DOT_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT_DOT
	 * @see #BORDER_SLANTED_DASH_DOT
	 */
	public void setBorderLeft(short border) {
		this.borderLeft = border;
	}

	private short borderLeft = BORDER_NONE;

	/**
	 * get the type of border to use for the left border of the cell
	 * 
	 * @return border type
	 * @see #BORDER_NONE
	 * @see #BORDER_THIN
	 * @see #BORDER_MEDIUM
	 * @see #BORDER_DASHED
	 * @see #BORDER_DOTTED
	 * @see #BORDER_THICK
	 * @see #BORDER_DOUBLE
	 * @see #BORDER_HAIR
	 * @see #BORDER_MEDIUM_DASHED
	 * @see #BORDER_DASH_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT
	 * @see #BORDER_DASH_DOT_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT_DOT
	 * @see #BORDER_SLANTED_DASH_DOT
	 */
	public short getBorderLeft() {
		return this.borderLeft;
	}

	/**
	 * set the type of border to use for the right border of the cell
	 * 
	 * @param border
	 *            type
	 * @see #BORDER_NONE
	 * @see #BORDER_THIN
	 * @see #BORDER_MEDIUM
	 * @see #BORDER_DASHED
	 * @see #BORDER_DOTTED
	 * @see #BORDER_THICK
	 * @see #BORDER_DOUBLE
	 * @see #BORDER_HAIR
	 * @see #BORDER_MEDIUM_DASHED
	 * @see #BORDER_DASH_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT
	 * @see #BORDER_DASH_DOT_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT_DOT
	 * @see #BORDER_SLANTED_DASH_DOT
	 */
	public void setBorderRight(short border) {
		this.borderRight = border;
	}

	private short borderRight = BORDER_NONE;

	/**
	 * get the type of border to use for the right border of the cell
	 * 
	 * @return border type
	 * @see #BORDER_NONE
	 * @see #BORDER_THIN
	 * @see #BORDER_MEDIUM
	 * @see #BORDER_DASHED
	 * @see #BORDER_DOTTED
	 * @see #BORDER_THICK
	 * @see #BORDER_DOUBLE
	 * @see #BORDER_HAIR
	 * @see #BORDER_MEDIUM_DASHED
	 * @see #BORDER_DASH_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT
	 * @see #BORDER_DASH_DOT_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT_DOT
	 * @see #BORDER_SLANTED_DASH_DOT
	 */
	public short getBorderRight() {
		return borderRight;
	}

	/**
	 * set the type of border to use for the top border of the cell
	 * 
	 * @param border
	 *            type
	 * @see #BORDER_NONE
	 * @see #BORDER_THIN
	 * @see #BORDER_MEDIUM
	 * @see #BORDER_DASHED
	 * @see #BORDER_DOTTED
	 * @see #BORDER_THICK
	 * @see #BORDER_DOUBLE
	 * @see #BORDER_HAIR
	 * @see #BORDER_MEDIUM_DASHED
	 * @see #BORDER_DASH_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT
	 * @see #BORDER_DASH_DOT_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT_DOT
	 * @see #BORDER_SLANTED_DASH_DOT
	 */
	public void setBorderTop(short border) {
		this.borderTop = border;
	}

	private short borderTop = BORDER_NONE;

	/**
	 * get the type of border to use for the top border of the cell
	 * 
	 * @return border type
	 * @see #BORDER_NONE
	 * @see #BORDER_THIN
	 * @see #BORDER_MEDIUM
	 * @see #BORDER_DASHED
	 * @see #BORDER_DOTTED
	 * @see #BORDER_THICK
	 * @see #BORDER_DOUBLE
	 * @see #BORDER_HAIR
	 * @see #BORDER_MEDIUM_DASHED
	 * @see #BORDER_DASH_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT
	 * @see #BORDER_DASH_DOT_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT_DOT
	 * @see #BORDER_SLANTED_DASH_DOT
	 */
	public short getBorderTop() {
		return borderTop;
	}

	/**
	 * set the type of border to use for the bottom border of the cell
	 * 
	 * @param border
	 *            type
	 * @see #BORDER_NONE
	 * @see #BORDER_THIN
	 * @see #BORDER_MEDIUM
	 * @see #BORDER_DASHED
	 * @see #BORDER_DOTTED
	 * @see #BORDER_THICK
	 * @see #BORDER_DOUBLE
	 * @see #BORDER_HAIR
	 * @see #BORDER_MEDIUM_DASHED
	 * @see #BORDER_DASH_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT
	 * @see #BORDER_DASH_DOT_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT_DOT
	 * @see #BORDER_SLANTED_DASH_DOT
	 */
	public void setBorderBottom(short border) {
		this.borderBottom = border;
	}

	private short borderBottom = BORDER_NONE;

	/**
	 * get the type of border to use for the bottom border of the cell
	 * 
	 * @return border type
	 * @see #BORDER_NONE
	 * @see #BORDER_THIN
	 * @see #BORDER_MEDIUM
	 * @see #BORDER_DASHED
	 * @see #BORDER_DOTTED
	 * @see #BORDER_THICK
	 * @see #BORDER_DOUBLE
	 * @see #BORDER_HAIR
	 * @see #BORDER_MEDIUM_DASHED
	 * @see #BORDER_DASH_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT
	 * @see #BORDER_DASH_DOT_DOT
	 * @see #BORDER_MEDIUM_DASH_DOT_DOT
	 * @see #BORDER_SLANTED_DASH_DOT
	 */
	public short getBorderBottom() {
		return this.borderBottom;
	}

	/**
	 * set the color to use for the left border
	 * 
	 * @param color
	 *            The index of the color definition
	 */
	public void setLeftBorderColor(short color) {
		this.leftBorderColor = color;
	}

	private short leftBorderColor = 0x8;

	/**
	 * get the color to use for the left border
	 */
	public short getLeftBorderColor() {
		return this.leftBorderColor;
	}

	/**
	 * set the color to use for the right border
	 * 
	 * @param color
	 *            The index of the color definition
	 */
	public void setRightBorderColor(short color) {
		this.rightBorderColor = color;
	}

	private short rightBorderColor = 0x8;

	/**
	 * get the color to use for the left border
	 * 
	 * @return the index of the color definition
	 */
	public short getRightBorderColor() {
		return this.rightBorderColor;
	}

	/**
	 * set the color to use for the top border
	 * 
	 * @param color
	 *            The index of the color definition
	 */
	public void setTopBorderColor(short color) {
		this.topBorderColor = color;
	}

	private short topBorderColor = 0x8;

	/**
	 * get the color to use for the top border
	 * 
	 * @return hhe index of the color definition
	 */
	public short getTopBorderColor() {
		return this.topBorderColor;
	}

	/**
	 * set the color to use for the bottom border
	 * 
	 * @param color
	 *            The index of the color definition
	 */
	public void setBottomBorderColor(short color) {
		this.bottomBorderColor = color;
	}

	private short bottomBorderColor = 0x8;

	/**
	 * get the color to use for the left border
	 * 
	 * @return the index of the color definition
	 */
	public short getBottomBorderColor() {
		return this.bottomBorderColor;
	}

	/**
	 * setting to one fills the cell with the foreground color... No idea about
	 * other values
	 *
	 * @see #NO_FILL
	 * @see #SOLID_FOREGROUND
	 * @see #FINE_DOTS
	 * @see #ALT_BARS
	 * @see #SPARSE_DOTS
	 * @see #THICK_HORZ_BANDS
	 * @see #THICK_VERT_BANDS
	 * @see #THICK_BACKWARD_DIAG
	 * @see #THICK_FORWARD_DIAG
	 * @see #BIG_SPOTS
	 * @see #BRICKS
	 * @see #THIN_HORZ_BANDS
	 * @see #THIN_VERT_BANDS
	 * @see #THIN_BACKWARD_DIAG
	 * @see #THIN_FORWARD_DIAG
	 * @see #SQUARES
	 * @see #DIAMONDS
	 *
	 * @param fp
	 *            fill pattern (set to 1 to fill w/foreground color)
	 */
	public void setFillPattern(short fp) {
		this.fp = fp;
	}

	private short fp = NO_FILL;

	/**
	 * get the fill pattern (??) - set to 1 to fill with foreground color
	 * 
	 * @return fill pattern
	 */

	public short getFillPattern() {
		return this.fp;
	}

	/**
	 * set the background fill color.
	 *
	 * @param bg
	 *            color
	 */

	public void setFillBackgroundColor(short bg) {
		this.fillBackgroundColor = bg;
	}

	private short fillBackgroundColor = 0;

	/**
	 * get the background fill color, if the fill is defined with an indexed
	 * color.
	 * 
	 * @return fill color index, or 0 if not indexed (XSSF only)
	 */
	public short getFillBackgroundColor() {
		return this.fillBackgroundColor;
	}

	/**
	 * set the foreground fill color <i>Note: Ensure Foreground color is set
	 * prior to background color.</i>
	 * 
	 * @param bg
	 *            color
	 */
	public void setFillForegroundColor(short bg) {
		this.fillForegroundColor = bg;
	}

	private short fillForegroundColor = 0;

	/**
	 * get the foreground fill color, if the fill is defined with an indexed
	 * color.
	 * 
	 * @return fill color, or 0 if not indexed (XSSF only)
	 */
	public short getFillForegroundColor() {
		return this.fillForegroundColor;
	}

	/**
	 * Controls if the Cell should be auto-sized to shrink to fit if the text is
	 * too long
	 */
	public void setShrinkToFit(boolean shrinkToFit) {
		this.shrinkToFit = shrinkToFit;
	}

	private boolean shrinkToFit = false;

	/**
	 * Should the Cell be auto-sized by Excel to shrink it to fit if this text
	 * is too long?
	 */
	public boolean getShrinkToFit() {
		return this.shrinkToFit;
	}

	public ColFont getFont() {
		return this.font;
	}

	private ColFont font;

	public void setFont(ColFont font) {
		this.font = font;
	}
}
