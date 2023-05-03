package com.indracompany.excellibs;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.indracompany.general.Log;

public abstract class WorkBookPlus {
	
	public static void copySheet(Workbook destWorkBook, Sheet srcSheet, Sheet destSheet, int scrRowNumFrom, int destRowNumFrom, boolean copyStyle) throws IOException{
	    int maxColumnNum = 0;     
	    Map<Integer, CellStyle> styleMap = (copyStyle) ? new HashMap<Integer, CellStyle>() : null;
	    
	    if (scrRowNumFrom <= 0) scrRowNumFrom = srcSheet.getFirstRowNum();
	    if (destRowNumFrom <= 0) destRowNumFrom = destSheet.getFirstRowNum();
	    
Log.print("Src row first ="+scrRowNumFrom);	    	  
Log.print("Src row last  ="+srcSheet.getLastRowNum());	    	  
	    for (int i = scrRowNumFrom, nRows=0; i <= srcSheet.getLastRowNum(); i++, nRows++) {  
Log.print("Src row "+i);	    	  
	      Row srcRow = srcSheet.getRow(i);     
	      Row destRow = destSheet.createRow(destRowNumFrom+nRows);    
	      if (srcRow != null) {  
Log.print("Copy row - Dest("+destRowNumFrom+"+"+nRows+") Src("+i+")");	    	  
	        copyRow(destWorkBook, srcSheet, destSheet, srcRow, destRow, styleMap);     
	        if (srcRow.getLastCellNum() > maxColumnNum) {     
	            maxColumnNum = srcRow.getLastCellNum();     
	        }
	      }     
	    }
	    // Change wordwarap style of cells 1,2 and 3 of col1
	    CellStyle style = destWorkBook.createCellStyle(); //Create new style
        style.setWrapText(false); //Set wordwrap
	    for (int i=0; i<3; i++) {
		    destSheet.getRow(i).getCell(0).setCellStyle(style);
	    }
	    for (int i = 0; i <= maxColumnNum; i++) {     
	      destSheet.setColumnWidth(i, srcSheet.getColumnWidth(i));     
	    }     
	}

	public static void copyRow(Workbook destWorkBook, Sheet srcSheet, Sheet destSheet, Row srcRow, Row destRow, Map<Integer, CellStyle> styleMap) throws IOException {
		Cell srcFromCol = srcRow.getCell(srcRow.getFirstCellNum());
		Cell srcToCol = srcRow.getCell(srcRow.getLastCellNum());
		Cell destCol = srcToCol;
		copyRowFromTo (destWorkBook, srcSheet, destSheet, srcRow, srcFromCol, srcToCol, destRow, destCol, styleMap);
	}

	public static void copyRowFromTo(Workbook destWorkBook, Sheet srcSheet, Sheet destSheet, Row srcRow, Cell srcFromCol, Cell srcToCol, Row destRow, Cell destCol, Map<Integer, CellStyle> styleMap) throws IOException {     
	    destRow.setHeight(srcRow.getHeight());
	    int scFrom = srcFromCol.getColumnIndex();
	    int scTo = srcToCol.getColumnIndex();
	    int destColNum = destCol.getColumnIndex();
Log.print("copyRowFromTo col (from,to)="+scFrom+","+scTo);
	    for (int j = scFrom, count = 0; j <= scTo; j++, count++) {     
	      Cell srcCell = srcRow.getCell(j);
	      Cell destCell = destRow.getCell(j);
	      if (srcCell != null) {     
	        if (destCell == null) {     
	          destCell = destRow.createCell(destColNum + count);
	        }     
	        copyCell(destWorkBook, srcCell, destCell, styleMap);
	      }
	    }                
	}
	
	public static void copyCell(Workbook destWorkBook, Cell srcCell, Cell destCell, Map<Integer, CellStyle> styleMap) {      
	    if(styleMap != null) {     
		      int stHashCode = srcCell.getCellStyle().hashCode();     
		      CellStyle destCellStyle = styleMap.get(stHashCode);     
		      if(destCellStyle == null){     
		        destCellStyle = destWorkBook.createCellStyle();     
		        destCellStyle.cloneStyleFrom(srcCell.getCellStyle());
		        styleMap.put(stHashCode, destCellStyle);     
		      }
		      destCell.setCellStyle(destCellStyle);
		}     
	    switch(srcCell.getCellType()) {     
	      case Cell.CELL_TYPE_STRING:     
	        destCell.setCellValue(srcCell.getRichStringCellValue());     
	        break;     
	      case Cell.CELL_TYPE_NUMERIC:     
	        destCell.setCellValue(srcCell.getNumericCellValue());     
	        break;     
	      case Cell.CELL_TYPE_BLANK:     
	        destCell.setCellType(Cell.CELL_TYPE_BLANK);     
	        break;     
	      case Cell.CELL_TYPE_BOOLEAN:     
	        destCell.setCellValue(srcCell.getBooleanCellValue());     
	        break;     
	      case Cell.CELL_TYPE_ERROR:     
	        destCell.setCellErrorValue(srcCell.getErrorCellValue());     
	        break;     
	      case Cell.CELL_TYPE_FORMULA:     
	        destCell.setCellFormula(srcCell.getCellFormula());     
	        break;     
	      default:     
	        break;     
	    }
	}		
}
