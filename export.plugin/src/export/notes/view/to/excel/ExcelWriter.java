package export.notes.view.to.excel;
/*
 * Copyright 2015
 * 
 * This file is part of Lotus Notes plugin for export Lotus Notes View into Microsoft Excel.
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); 
 * you may not use this file except in compliance with the License. 
 * You may obtain a copy of the License at:
 * 
 * http://www.apache.org/licenses/LICENSE-2.0 
 * 
 * Unless required by applicable law or agreed to in writing, software 
 * distributed under the License is distributed on an "AS IS" BASIS, 
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or 
 * implied. See the License for the specific language governing 
 * permissions and limitations under the License.
 */

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Vector;

import lotus.domino.DateTime;
import lotus.domino.NotesException;
import lotus.domino.RichTextStyle;
import lotus.domino.View;
import lotus.domino.ViewColumn;
import lotus.domino.ViewEntry;
import lotus.domino.ViewNavigator;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	XSSFWorkbook workbook;
	XSSFSheet sheet;

	Map<Integer, ViewColumn> headers = new HashMap<Integer, ViewColumn>();
	ArrayList<CellStyle> styles = new ArrayList<CellStyle>();

	public ExcelWriter() {
		workbook = new XSSFWorkbook();
	}

	public XSSFWorkbook getWorkbook() {
		return workbook;
	}

	public XSSFSheet createSheet(String name) {
		if (workbook == null) {
			workbook = new XSSFWorkbook();
		}
		if (sheet == null) {
			sheet = workbook.createSheet(name.replaceAll("\\\\", "-").trim());//$NON-NLS-1$ //$NON-NLS-2$
		}
		return sheet;
	}

	public void cerateRow(int rowNum, List<Object> entryValues) throws NotesException {
		ArrayList<Object> values = new ArrayList<Object>();

		for (Entry<Integer, ViewColumn> entry : headers.entrySet()) {
			Object object = entryValues.get(entry.getKey());
			if (object.getClass().getName().contains("Vector")) { //$NON-NLS-1$				
				values.add(getVectorString(object));
			} else {
				values.add(object);
			}
		}

		// write the row into excel table
		Row row = sheet.createRow(rowNum);
		for (int i = 0; i < values.size(); i++) {
			Cell cell = row.createCell(i);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			Object object = values.get(i);
			String name = object.getClass().getSimpleName();
			if (name.contains("String")) { //$NON-NLS-1$											
				cell.setCellValue((String) object);
			} else if (name.contains("Double")) { //$NON-NLS-1$
				cell.setCellValue((Double) object);
			} else if (name.contains("DateTime")) { //$NON-NLS-1$				
				cell.setCellValue(((DateTime) values.get(i)).toJavaDate());
			} else {
				cell.setCellValue(object.toString());
			}
			cell.setCellStyle(styles.get(i));
		}

	}

	public void setAutoSizeColumns() {
		for (int x = 0; x < headers.size(); x++) {
			sheet.autoSizeColumn(x);
			if (sheet.getColumnWidth(x) > 25000) {
				sheet.setColumnWidth(x, 25000);
			}
		}
	}

	private String getVectorString(Object object) {
		String s = object.toString();
		// remove brackets from Vector values
		if (s.startsWith("[")) { //$NON-NLS-1$
			s = s.substring(1);
		}
		if (s.endsWith("]")) { //$NON-NLS-1$
			s = s.substring(0, s.length() - 1);
		}
		return s;
	}

	// create header of table
	@SuppressWarnings("unchecked")
	public void createTableHeader(View view) throws NotesException {
		Vector<ViewColumn> columns = view.getColumns();
		// offset column
		int offset = 0;
		for (int x = 0; x < columns.size(); x++) {
			ViewColumn column = columns.get(x);
			if (column.isConstant()) {
				offset++;
			} else if (column.isFormula()) {
				// A column value (ViewEntry.getColumnValues()) is not
				// returned if it is determined by a constant. Check it.
				String formula = column.getFormula();
				if (formula == null) {
					offset++;
				} else {
					// empty string
					formula = formula.replaceAll("\"", "").trim();
					// some whitespaces
					if (StringUtils.isBlank(formula)) {
						offset++;
					}
				}
			}
			// hidden and icons columns not will be use
			if (!column.isHidden() && !column.isIcon() && !column.isConstant()) {
				String s = columns.get(x).getTitle();
				if (s == null || "".equals(s)) { //$NON-NLS-1$
					s = " "; //$NON-NLS-1$
				}
				int position = x - offset;
				headers.put(position, column);
				ViewNavigator nav = view.createViewNav();
				ViewEntry entry = nav.getFirst();
				while (!entry.isDocument()) {
					entry = nav.getNext();
				}
				createCellStyle(position, column, entry);
			}
		}

		// column indexes
		int idy = 0;
		// Generate column headings
		Cell c = null;
		if (sheet == null) {
			SimpleDateFormat sf = new SimpleDateFormat("dd.MM.yyyy HH:mm"); //$NON-NLS-1$
			sheet = createSheet(Messages.ExportAction_10 + " " + sf.format(new Date())); 
		}
		Row row = sheet.createRow(0);

		Font fontBold = workbook.createFont();
		fontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFont(fontBold);

		for (Entry<Integer, ViewColumn> entry : headers.entrySet()) {
			ViewColumn column = entry.getValue();
			c = row.createCell(idy++);
			c.setCellValue(column.getTitle());
			c.setCellStyle(cellStyle);
		}

		sheet.createFreezePane(0, 1, 0, 1);
	}

	private void createCellStyle(int position, ViewColumn column, ViewEntry entry) throws NotesException {
		CellStyle cellStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		if (column.isFontBold()) {
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		}
		font.setItalic(column.isFontItalic());
		switch (column.getFontColor()) {
		case RichTextStyle.COLOR_BLACK:
			font.setColor(HSSFColor.BLACK.index);
			break;
		case RichTextStyle.COLOR_BLUE:
			font.setColor(HSSFColor.BLUE.index);
			break;
		case RichTextStyle.COLOR_CYAN:
			font.setColor(HSSFColor.CORAL.index);
			break;
		case RichTextStyle.COLOR_DARK_BLUE:
			font.setColor(HSSFColor.DARK_BLUE.index);
			break;
		case RichTextStyle.COLOR_DARK_CYAN:
			font.setColor(HSSFColor.DARK_GREEN.index);
			break;
		case RichTextStyle.COLOR_DARK_GREEN:
			font.setColor(HSSFColor.DARK_GREEN.index);
			break;
		case RichTextStyle.COLOR_DARK_MAGENTA:
			font.setColor(HSSFColor.VIOLET.index);
			break;
		case RichTextStyle.COLOR_DARK_RED:
			font.setColor(HSSFColor.DARK_RED.index);
			break;
		case RichTextStyle.COLOR_DARK_YELLOW:
			font.setColor(HSSFColor.DARK_YELLOW.index);
			break;
		case RichTextStyle.COLOR_GRAY:
			font.setColor(HSSFColor.GREY_80_PERCENT.index);
			break;
		case RichTextStyle.COLOR_GREEN:
			font.setColor(HSSFColor.GREEN.index);
			break;
		case RichTextStyle.COLOR_LIGHT_GRAY:
			font.setColor(HSSFColor.GREY_50_PERCENT.index);
			break;
		case RichTextStyle.COLOR_MAGENTA:
			font.setColor(HSSFColor.VIOLET.index);
			break;
		case RichTextStyle.COLOR_RED:
			font.setColor(HSSFColor.RED.index);
			break;
		case RichTextStyle.COLOR_WHITE:
			font.setColor(HSSFColor.BLACK.index);
			break;
		case RichTextStyle.COLOR_YELLOW:
			font.setColor(HSSFColor.YELLOW.index);
			break;
		default:
			break;
		}

		cellStyle.setFont(font);

		switch (column.getAlignment()) {
		case ViewColumn.ALIGN_CENTER:
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
			break;
		case ViewColumn.ALIGN_LEFT:
			cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
			break;
		case ViewColumn.ALIGN_RIGHT:
			cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);
			break;
		default:
			break;
		}

		@SuppressWarnings("unchecked")
		Vector<Object> values = entry.getColumnValues();
		Object value = values.get(position);
		String name = value.getClass().getSimpleName();
		short format = 0;
		if (name.contains("Double")) { //$NON-NLS-1$
			XSSFDataFormat fmt = (XSSFDataFormat) workbook.createDataFormat();
			switch (column.getNumberFormat()) {
			case ViewColumn.FMT_CURRENCY:
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(6));
				break;
			case ViewColumn.FMT_FIXED:
				String zero = "0"; //$NON-NLS-1$
				String fixedFormat = "#0"; //$NON-NLS-1$
				int digits = column.getNumberDigits();
				if (digits > 0) {
					String n = StringUtils.repeat(zero, digits);
					fixedFormat = fixedFormat + "." + n;
				}
				format = fmt.getFormat(fixedFormat);
				break;
			default:
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(1));
				break;
			}
		} else if (name.contains("DateTime")) { //$NON-NLS-1$							
			XSSFDataFormat fmt = (XSSFDataFormat) workbook.createDataFormat();
			switch (column.getTimeDateFmt()) {
			case ViewColumn.FMT_DATE:
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(0xe));
				break;
			case ViewColumn.FMT_DATETIME:
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(0x16));
				break;
			case ViewColumn.FMT_TIME:
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(0x15));
				break;
			default:
				format = fmt.getFormat(BuiltinFormats.getBuiltinFormat(0xe));
				break;
			}			
		}
		cellStyle.setDataFormat(format);
		styles.add(cellStyle);
	}
}
