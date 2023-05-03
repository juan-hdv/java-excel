package excelfilter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.indracompany.excellibs.WorkBookPlus;
import com.indracompany.general.Log;

public class ExcelFilter {

	public int DeleteRowsInFilter (File srcFile, File destFile, File filterFile, int srcSheetNum,  int srcRowNum, int srcColNum, int filterSheetNum, int filterRowNum, int filterColNum) throws InvalidFormatException, IOException {
		return DeleteRows_fromFilterFile (srcFile, destFile, filterFile, srcSheetNum, srcRowNum, srcColNum, filterSheetNum, filterRowNum, filterColNum, true, true);
	}

	public int DeleteRowsNotInFilter (File srcFile, File destFile, File filterFile, int srcSheetNum,  int srcRowNum, int srcColNum, int filterSheetNum, int filterRowNum, int filterColNum) throws InvalidFormatException, IOException {
		return DeleteRows_fromFilterFile (srcFile, destFile, filterFile, srcSheetNum, srcRowNum, srcColNum, filterSheetNum, filterRowNum, filterColNum, false, true);
	}

	public int DeleteRowsContainingPhrases (File srcFile, File destFile, String[] phrases, int srcSheetNum,  int srcRowNum, int srcColNum) throws InvalidFormatException, IOException {
		return DeleteRows_fromFilterPhrases (srcFile, destFile , phrases, srcSheetNum,  srcRowNum, srcColNum, true, true);
	}

	public int DeleteRowsNotContainingPhrases (File srcFile, File destFile, String[] phrases, int srcSheetNum,  int srcRowNum, int srcColNum) throws InvalidFormatException, IOException {
		return DeleteRows_fromFilterPhrases (srcFile, destFile , phrases, srcSheetNum,  srcRowNum, srcColNum, false, true);
	}

	/*
	 * Creates a copy of the source file for working on it
	 * Process target sheet, filtering rows
	 * Save to new file
	 */
	public int DeleteRows_fromFilterFile (File srcFile, File destFile, File filterFile, int srcSheetNum,  int srcRowNum, int srcColNum, int filterSheetNum, int filterRowNum, int filterColNum, boolean elementsInFilter, boolean deleteEmptyElements) throws InvalidFormatException, IOException {
Log.print("DeleteRows_fromFilterFile Sheet,Row,Col = "+srcSheetNum+","+srcRowNum+","+srcColNum);
		
		// Opens a copy of source file as a template for working on it
		Workbook destBook = WorkbookFactory.create(srcFile);
		Workbook srcBook = WorkbookFactory.create(srcFile);
		Workbook filterBook = WorkbookFactory.create(filterFile);
		
		Sheet srcSheet = srcBook.getSheetAt(srcSheetNum);
		// Delete the target sheet
		destBook.removeSheetAt(srcSheetNum);
		// Create a new Sheet in destBook and set order
		String sheetName = srcSheet.getSheetName();
		Sheet destSheet = destBook.createSheet(sheetName);
		destBook.setSheetOrder(sheetName,srcSheetNum);
		// Get filter sheet
		Sheet filterSheet = filterBook.getSheetAt(filterSheetNum);
		
		boolean foundElement; 
		int filterCount = 0;
		ArrayList<Integer> rowList = new ArrayList<Integer>();
		
		int srcLastRowNum = srcSheet.getLastRowNum(); // Rows from 0
		int filterLastRowNum = filterSheet.getLastRowNum(); // Rows from 0
		for (int i=srcRowNum; i <= srcLastRowNum; i++) {
			String srcValue = cell2string (srcSheet.getRow(i).getCell(srcColNum));
			foundElement = false; 
			// Loop through the filter file
			for (int j=filterRowNum; j <= filterLastRowNum; j++) {
				Cell filterCell = filterSheet.getRow(j).getCell(filterColNum);
				String filterValue = cell2string (filterCell);
				if (srcValue.equalsIgnoreCase(filterValue)) {
Log.print("Src("+i+","+srcColNum+") == Filter("+j+","+filterColNum+") == "+filterValue);			
					foundElement = true;
					break;
				}
			}
			if ((deleteEmptyElements && srcValue.isEmpty()) || // Delete if Empty
			    (elementsInFilter && foundElement) ||    // Delete if "elementsInFilter"
				(!elementsInFilter && !foundElement)) {  // Delete if NOT "elementsInFilter"
				// Delete if found
				rowList.add(i);
Log.print("A borrar "+i);
			}
		} // End for
		// Now filter rows
		// PRE: rowList = list of all the rows that must NOT be present in destination file 
		filterCount = copyRows2NewSheet (destBook, srcSheet, destSheet, 0, 0, rowList, false);
		
	    try {
			FileOutputStream out = new FileOutputStream(destFile);
			destBook.write(out);
			out.close();
	    } catch (Exception e) {
	    	e.printStackTrace();
		}
	    return filterCount;
	} // End filterDeleteRows_fromFilterFile

	public int DeleteRows_fromFilterPhrases (File srcFile, File destFile, String[] phrases, int srcSheetNum,  int srcRowNum, int srcColNum, boolean elementMatchPhrase, boolean deleteEmptyElements) throws InvalidFormatException, IOException {
Log.print("DeleteRows_fromFilterPhrase Sheet,Row,Col = "+srcSheetNum+","+srcRowNum+","+srcColNum);			
Log.print ("Phrases = " + StringUtils.join (phrases,","));

		// Opens a copy of source file as a template for working on it
		Workbook destBook = WorkbookFactory.create(srcFile);
		Workbook srcBook = WorkbookFactory.create(srcFile);
		
		Sheet srcSheet = srcBook.getSheetAt(srcSheetNum);
		// Delete the target sheet
		destBook.removeSheetAt(srcSheetNum);
		// Create a new Sheet in destBook and set order
		String sheetName = srcSheet.getSheetName();
		Sheet destSheet = destBook.createSheet(sheetName);
		destBook.setSheetOrder(sheetName,srcSheetNum);
			
		boolean foundElement; 
		int filterCount = 0;
		ArrayList<Integer> rowList = new ArrayList<Integer>();
		
		// Process srcSheet--- 
		int srcLastRowNum = srcSheet.getLastRowNum(); // Rows from 0
		for (int i=srcRowNum; i <= srcLastRowNum; i++) {
			String srcValue = cell2string (srcSheet.getRow(i).getCell(srcColNum)).toLowerCase();
			foundElement = false;
			String tmp = "";
			for (String p : phrases) {
				foundElement = srcValue.startsWith(p);
				if (foundElement) {
					tmp = p;
					break;
				}
			}
Log.print("Row "+i+". Element="+srcValue+(foundElement?"==":"<>")+(tmp.isEmpty()?"ALL PHRASES":tmp));
			if ((deleteEmptyElements && srcValue.isEmpty()) || // Delete if Empty
			    (elementMatchPhrase && foundElement) ||    // Delete if "elementMatchPhrase"
				(!elementMatchPhrase && !foundElement)) {  // Delete if NOT "elementMatchPhrase"
				// Delete if found
				rowList.add(i);
Log.print("To delete row "+i+". Element="+srcValue+". Found="+(foundElement?"true":"false"));
			}
		} // End for
		// Now filter rows
		// PRE: rowList = list of all the rows that must NOT be present in destination file 
		filterCount = copyRows2NewSheet (destBook, srcSheet, destSheet, 0, 0, rowList, false);
		
		// WorkBookPlus.copySheet(destBook, srcSheet, destSheet, 0, 0, true);			
	    try {
			FileOutputStream out = new FileOutputStream(destFile);
			destBook.write(out);
			out.close();
	    } catch (Exception e) {
	    	e.printStackTrace();
		}
	    return filterCount;
	} // End filterDeleteRows_fromFilterPhrase
	
	private String cell2string(Cell cell) {
		if (cell == null)
			return "";
		int type = cell.getCellType();
		Object result = null;
		
		if (type == Cell.CELL_TYPE_FORMULA) {
		    // result = cell.getCellFormula();
			type = cell.getCachedFormulaResultType();
		}
		switch (type) {
		case Cell.CELL_TYPE_NUMERIC:
			result = cell.getNumericCellValue();
			break;
		case Cell.CELL_TYPE_STRING:
			result = cell.getRichStringCellValue().getString().trim();
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			result = cell.getBooleanCellValue();
			break;
		case Cell.CELL_TYPE_ERROR:
			result = cell.getErrorCellValue();
			break;
		case Cell.CELL_TYPE_BLANK:
			// Blank cell
		case Cell.CELL_TYPE_FORMULA:
			result = "";
			// Will never occur
			break;
		default:	
			throw new RuntimeException ("Qué celda tan extraña");
		}
		return result.toString();
	}

	/*
	 * Copy all the rows @in (In/Not in) rowList
	 * from @destBook - @srcSheet begining at @srcRow 
	 * to   @destSheet begining at @destRow
	 */
	public int copyRows2NewSheet (Workbook destBook, Sheet srcSheet, Sheet destSheet, int srcRowNum, int destRowNum, ArrayList<Integer> rowList, boolean in) throws IOException {
	// (Sheet sheet, int fromRow, ArrayList<Integer> rowList) {
		Map<Integer, CellStyle> styleMap = new HashMap<Integer, CellStyle>();

		// Create destRow rows at the begining of destSheet
		int copiedRows=0;
		for (int i=0; i < destRowNum; i++) destSheet.createRow(copiedRows++);
		
		int maxColumnNum = 0;
		int srcLastRowNum=srcSheet.getLastRowNum();
		for (int i=srcRowNum; i <= srcLastRowNum; i++) {
			boolean rowInList = rowList.contains(i);
			if ((!in && !rowInList) || (in && rowInList)) { // => Row i must be copied
				Row srcRowObject = srcSheet.getRow(i);
				Row destRowObject = destSheet.createRow(copiedRows++);				
				if (srcRowObject != null) {  
					WorkBookPlus.copyRow(destBook, srcSheet, destSheet, srcRowObject, destRowObject, styleMap);
					if (srcRowObject.getLastCellNum() > maxColumnNum) {     
			            maxColumnNum = srcRowObject.getLastCellNum();     
			        }     					
Log.print("Copying row " + i + ". Displaying Col 0:" + cell2string(destRowObject.getCell(0)));
				}
			}
		} // End for
	    // Change wordwarap style of cells 1,2 and 3 of col1
	    CellStyle style = destBook.createCellStyle(); //Create new style
        style.setWrapText(false); //Set wordwrap
	    for (int i=0; i<3; i++) {
		    destSheet.getRow(i).getCell(0).setCellStyle(style);
	    }
	    // Set col width for all columns of destSheet
	    for (int i = 0; i <= maxColumnNum; i++) {     
	      destSheet.setColumnWidth(i, srcSheet.getColumnWidth(i));     
	    }     
		return copiedRows;
	}

	/**
	 * Remove a row by index
	 * @param sheet 
	 * @param rowIndex a 0 based index of removing row
	 */
	public static void removeRow(Sheet sheet, int rowIndex) {
	    int lastRowNum=sheet.getLastRowNum();
	    if(rowIndex>=0 && rowIndex<lastRowNum){
	        sheet.shiftRows(rowIndex+1,lastRowNum,-1,true,true);
	    }
        Row r=sheet.getRow(lastRowNum);
        if(r!=null) sheet.removeRow(r);
	}
}

