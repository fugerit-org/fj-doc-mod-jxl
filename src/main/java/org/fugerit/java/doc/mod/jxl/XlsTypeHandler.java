package org.fugerit.java.doc.mod.jxl;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Iterator;

import org.fugerit.java.core.lang.helpers.BooleanUtils;
import org.fugerit.java.doc.base.config.DocConfig;
import org.fugerit.java.doc.base.config.DocException;
import org.fugerit.java.doc.base.config.DocInput;
import org.fugerit.java.doc.base.config.DocOutput;
import org.fugerit.java.doc.base.config.DocTypeHandler;
import org.fugerit.java.doc.base.config.DocTypeHandlerDefault;
import org.fugerit.java.doc.base.model.DocBase;
import org.fugerit.java.doc.base.model.DocBorders;
import org.fugerit.java.doc.base.model.DocCell;
import org.fugerit.java.doc.base.model.DocElement;
import org.fugerit.java.doc.base.model.DocPara;
import org.fugerit.java.doc.base.model.DocPhrase;
import org.fugerit.java.doc.base.model.DocRow;
import org.fugerit.java.doc.base.model.DocTable;
import org.fugerit.java.doc.base.typehelper.excel.ExcelHelperConsts;
import org.fugerit.java.doc.base.typehelper.excel.TableMatrix;
import org.fugerit.java.doc.base.typehelper.generic.FormatTypeConsts;
import org.fugerit.java.doc.base.xml.DocModelUtils;

import jxl.CellView;
import jxl.Workbook;
import jxl.biff.DisplayFormat;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.format.RGB;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.NumberFormats;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import lombok.AllArgsConstructor;
import lombok.Getter;

public class XlsTypeHandler extends DocTypeHandlerDefault {
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 8856614360877809158L;

	public static final DocTypeHandler HANDLER = new XlsTypeHandler();
	
	public static final String MODULE = "jxl";
	
	public XlsTypeHandler() {
		super( DocConfig.TYPE_XLS, MODULE );
	}

	@Override
	public void handle(DocInput docInput, DocOutput docOutput) throws Exception {
		DocBase docBase = docInput.getDoc();
		OutputStream outputStream = docOutput.getOs();
		String excelTemplate = docBase.getInfo().getProperty( ExcelHelperConsts.PROP_XLS_TEMPLATE );
		Workbook templateXls = null;
		if ( excelTemplate != null ) {
			File templateFile = new File( excelTemplate );
			try ( InputStream is = new FileInputStream( templateFile ) ) {
				templateXls = Workbook.getWorkbook( is );	
			}
		}			
		XlsTypeHandler.handleDoc( docBase , outputStream, templateXls );
	}

	private static int rgbDiff( RGB rgb1, RGB rgb2 ) {
		int result = 0;
		result+= ( Math.abs( rgb1.getRed() - rgb2.getRed() ) );
		result+= ( Math.abs( rgb1.getBlue() - rgb2.getBlue() ) );
		result+= ( Math.abs( rgb1.getGreen() - rgb2.getGreen() ) );
		return result;
	}
	
	// find the closest jxl colour to a give awt colour
	public static Colour closestColor( Color awtColor ) {
		Colour cc = null;
		Colour[] c = Colour.getAllColours();
		RGB rgbBase = new RGB( awtColor.getRed(), awtColor.getGreen(), awtColor.getBlue() );
		int rgbDiff = 10000000;
		for ( int a=0; a<c.length; a++ ) {
			Colour t = c[a];
			RGB rgb = t.getDefaultRGB();
			int currDiff = rgbDiff( rgb, rgbBase );
			if ( rgbDiff > currDiff  ) {
				rgbDiff = currDiff;
				cc = t;
			}
		}
		return cc;
	}
	
	private static BorderLineStyle getBorderStyle( int borderWidth ) {
		BorderLineStyle bls = BorderLineStyle.THIN;
		if ( borderWidth == 0 ) {
			bls = BorderLineStyle.NONE;
		} else if ( borderWidth > 5 ) {
			bls = BorderLineStyle.THICK;
		} else if ( borderWidth > 3 ) {
			bls = BorderLineStyle.MEDIUM;			
		}		
		return bls;
	}
	
	private static void handleBorders(TableMatrix matrix, int rn, int cn, DocCell cell, WritableCellFormat cf) throws CloneNotSupportedException, WriteException {
		// borders
		DocBorders borders = matrix.getBorders(rn, cn);
		cf.setBorder(Border.LEFT, getBorderStyle(borders.getBorderWidthLeft()));
		cf.setBorder(Border.RIGHT, getBorderStyle(borders.getBorderWidthRight()));
		cf.setBorder(Border.BOTTOM, getBorderStyle(borders.getBorderWidthBottom()));
		cf.setBorder(Border.TOP, getBorderStyle(borders.getBorderWidthTop()));
		if (cell != null) {
			// alignment
			if (cell.getAlign() == DocPara.ALIGN_CENTER) {
				cf.setAlignment(Alignment.CENTRE);
			} else if (cell.getAlign() == DocPara.ALIGN_RIGHT) {
				cf.setAlignment(Alignment.RIGHT);
			} else if (cell.getAlign() == DocPara.ALIGN_LEFT) {
				cf.setAlignment(Alignment.LEFT);
			}
			// vertical alignment
			if (cell.getValign() == DocPara.ALIGN_MIDDLE) {
				cf.setVerticalAlignment(VerticalAlignment.CENTRE);
			} else if (cell.getValign() == DocPara.ALIGN_BOTTOM) {
				cf.setVerticalAlignment(VerticalAlignment.BOTTOM);
			} else if (cell.getValign() == DocPara.ALIGN_TOP) {
				cf.setVerticalAlignment(VerticalAlignment.TOP);
			}
		}
	}
	
	private static WritableCellFormat handleParaStyle( DocPara currentePara, WritableCellFormat cf, DisplayFormat df) throws WriteException {
		if ( currentePara != null ) {
			Font f = cf.getFont();
			WritableFont wf = new WritableFont( f );
			if ( currentePara.getStyle() == DocPara.STYLE_BOLD ) {
				wf.setBoldStyle( WritableFont.BOLD );
			} else if ( currentePara.getStyle() == DocPara.STYLE_ITALIC ) {
				wf.setItalic( true );
			} else if ( currentePara.getStyle() == DocPara.STYLE_BOLDITALIC ) {	
				wf.setBoldStyle( WritableFont.BOLD );
				wf.setItalic( true );
			} else if ( currentePara.getStyle() == DocPara.STYLE_UNDERLINE ) {
				wf.setUnderlineStyle( UnderlineStyle.SINGLE );
			}
			if ( df != null ) {
				cf = new WritableCellFormat( wf, df );	
			} else {
				cf = new WritableCellFormat( wf );
			}
		}
		return cf;
	}
	
	private static WritableCellFormat handleFormat( WritableCellFormat cf, DocCell parent, DisplayFormat df, DocPara currentePara ) throws WriteException {
		// must go first as it has the chance to change the cell format
		if ( parent.getForeColor() != null ) {
			Font f = cf.getFont();
			WritableFont wf = new WritableFont( f );
			wf.setColour( ( closestColor( DocModelUtils.parseHtmlColor( parent.getForeColor() ) ) ) );
			if ( df != null ) {
				cf = new WritableCellFormat( wf, df );	
			} else {
				cf = new WritableCellFormat( wf );
			}
		}	
		// style
		cf = handleParaStyle(currentePara, cf, df);
		return cf;
	}
	
	private static WritableCellFormat handleParent( JxlContext context, WritableCellFormat cf, DocCell cell, DocCell parent, int rn, int cn ) throws WriteException, CloneNotSupportedException {
		if ( parent != null && !context.isIgnoreFormat() ) {
			// back color
			if ( parent.getBackColor() != null) {
				cf.setBackground( closestColor( DocModelUtils.parseHtmlColor( parent.getBackColor() ) ) );
			}
			handleBorders( context.getTableMatrix(), rn, cn, cell, cf);
		}
		return cf;
	}
	
	private static DisplayFormat handleFormat( String format, String type ) {
		DisplayFormat df = null;
		if ( format != null && FormatTypeConsts.TYPE_NUMBER.equalsIgnoreCase( type ) ) {
			if ( "float".equalsIgnoreCase( format ) ) {
				df = NumberFormats.FLOAT;
			} else {
				df = new NumberFormat( format, NumberFormat.COMPLEX_FORMAT);
			}
		}
		return df;
	}
	
	private static void handleCell( JxlContext context, int rn, int cn) throws WriteException, CloneNotSupportedException {
		String type = null;
		String format = null;
		DocCell cell = context.getTableMatrix().getCell( rn, cn );
		DocCell parent = context.getTableMatrix().getParent( rn, cn );
		String text = "";
		DocPara currentePara = null;
		if ( cell != null ) {
			Iterator<DocElement> it1 = cell.docElements();
			DocElement current = it1.next();
			if ( current instanceof DocPara ) {
				currentePara = ((DocPara)current);
				text = currentePara.getText();
				type = currentePara.getType();
				format = currentePara.getFormat();
			} else if ( current instanceof DocPhrase ) {
				DocPhrase currentPhrase = ((DocPhrase)current);
				text = currentPhrase.getText();
			} else {
				text = "Type not handled ("+current+")";
			}
		}
		// in case of number check for format string
		DisplayFormat df = handleFormat(format, type);
		WritableCellFormat cf = ( df == null ) ? new WritableCellFormat() : new WritableCellFormat( df );
		cf = handleFormat(cf, parent, df, currentePara);
		cf = handleParent( context, cf, cell, parent, rn, cn );
		WritableCell current = null;
		if ( FormatTypeConsts.TYPE_NUMBER.equalsIgnoreCase( type ) ) {
			BigDecimal bd = new BigDecimal( text );
			current = new Number( cn, rn,  bd.doubleValue(), cf );
		} else {
			current = new Label( cn, rn, text, cf );
		}
		context.getSheet().addCell( current );	
	}
	
	private static TableMatrix handleMatrix( DocTable table, boolean ignoreFormat, WritableSheet dati ) throws WriteException, CloneNotSupportedException {
		TableMatrix matrix = new TableMatrix( table.containerSize() , table.getColumns() );
		JxlContext context = new JxlContext(matrix, ignoreFormat, dati);
		Iterator<DocElement> rows = table.docElements();
		while ( rows.hasNext() ) {
			DocRow row = (DocRow)rows.next();
			Iterator<DocElement> cells = row.docElements();
			while ( cells.hasNext() ) {
				DocCell cell = (DocCell)cells.next();
				matrix.setNext( cell, cell.getRSpan() , cell.getCSpan() );
			}
		}
		for ( int rn=0; rn<matrix.getRowCount(); rn++ ) {
			for ( int cn=0; cn<matrix.getColumnCount(); cn++ ) {
				handleCell(context, rn, cn);
			}
		}
		return matrix;
	}

	private static void handleMerge( DocTable table, boolean ignoreFormat, WritableSheet sheet ) throws WriteException, CloneNotSupportedException {
		TableMatrix matrix = handleMatrix(table, ignoreFormat, sheet);
		for ( int rn=0; rn<matrix.getRowCount(); rn++ ) {
			for ( int cn=0; cn<matrix.getColumnCount(); cn++ ) {
				DocCell cell = matrix.getCell( rn, cn );
				if ( cell != null ) {
					int rs = cell.getRSpan()-1;
					int cs = cell.getCSpan()-1;
					if ( rs != 0 || cs != 0 ) {
						sheet.mergeCells( cn, rn, cn+cs, rn+rs );	
					}
				}
			}
		}
	}

	public static void handleDoc( DocBase docBase, OutputStream os, Workbook templateXls ) throws DocException {
		try {
			String excelTableId = docBase.getInfo().getProperty( ExcelHelperConsts.PROP_XLS_TABLE_ID );
			String[] excelTableSheet = excelTableId.split( ";" );
			WritableWorkbook outputXls = null;
			if ( templateXls == null ) {
				outputXls = Workbook.createWorkbook( os );
			} else {
				outputXls = Workbook.createWorkbook( os, templateXls );
			}	
			boolean ignoreFormat = BooleanUtils.isTrue( docBase.getInfo().getProperty( ExcelHelperConsts.PROP_XLS_IGNORE_FORMAT ) );
			for ( int k=0; k<excelTableSheet.length; k++ ) {
				String[] currentSheetData = excelTableSheet[k].split( "=" );
				String sheetId = currentSheetData[0];
				String sheetName = currentSheetData[1];
				DocTable table = (DocTable)docBase.getElementById( sheetId );
				
				WritableSheet sheet = null;
				if ( templateXls == null ) {
					sheet = outputXls.createSheet( sheetName , k );
				} else {
					sheet = outputXls.getSheet( k );
					sheet.setName( sheetName );
				}
				
				int[] colW = table.getColWithds();
				for ( int i=0; i<colW.length; i++ ) {
					CellView cw = new CellView( sheet.getColumnView( i ) );
					int mul = Integer.parseInt( docBase.getInfo().getProperty( ExcelHelperConsts.PROP_XLS_WIDTH_MULTIPLIER, ExcelHelperConsts.PROP_XLS_WIDTH_MULTIPLIER_DEFAULT ) );
					cw.setSize( colW[i]* mul );
					sheet.setColumnView( i , cw );
				}
				
				handleMerge(table, ignoreFormat, sheet);
				
			}
			outputXls.write();
			outputXls.close();
		} catch (Exception e) {
			throw DocException.convertExMethod( "handleDoc", e );
		}
	}

}

@AllArgsConstructor
class JxlContext {
	
	@Getter private TableMatrix tableMatrix;
	
	@Getter private boolean ignoreFormat;
	
	@Getter private WritableSheet sheet;
	
}