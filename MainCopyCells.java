import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.indracompany.excellibs.WorkBookPlus;
import com.indracompany.general.Log;

public class MainCopyCells {
	
	static String CMD_USE = 
			"USE:\n" +
			"  copy_cells --srcArea=<sheet,RowFrom,RowTo,ColFrom,ColTo> --destCell=<sheet,RowTo,ColTo> <srcFilePath1> <srcFilePath2> <destFilePath>\n" +
			"* Copy <srcArea> to <destCell> from <srcFilePath2> to <srcFilePath2> resulting in <destFuilePath>\n" +
			"  Row = -1 means Last Row; Col = -1 means Last Column";
	
	final static String REGEX_PATH = "([a-zA-Z]:)?([\\\\]{0,2}[a-zA-Z0-9 _.-]+)+[\\\\]{0,2}";

	private static Scanner scanner2;
	
	/*
	 * Copy cells area from srcFile.sheet(?) to destFile.sheet(?)
	 */
	public static void main(String[] args) throws Exception {

		// Log.setup(true,"myapp.log");
		Log.setup(false,"");
		
		String tmpArg = "";
		String destPath="", srcPath1="", srcPath2 = "tmp.xls";
		int srcSheetNum=0,srcRowFrom=0,srcRowTo=-1,srcColFrom=0, srcColTo=-1;
		int destSheetNum=0,destRowTo=0,destColTo=0;
		String a[]; // Temp array for splitting strings
		String params; // Comma separated string of sheet,row,col...
		int numParms;
		Scanner sc; 
		Pattern regex;
		Matcher m;
		int argNumber = 0;
		
		int nArgs = args.length;
		if (nArgs < 5 || nArgs > 5) {
			print_error("Bad number of arguments.");
			return;
		}
		
		// Parse --srcArea=<sheet,RowFrom,RowTo,ColFrom,ColTo>
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.startsWith("--srcarea=")) {
			print_error("Missing argument "+(argNumber+1)+":  --srcArea=<sheet,RowFrom,RowTo,ColFrom,ColTo> => "+tmpArg);
			return;
		}	
		a = tmpArg.split("=");
		params = a[1];
		numParms=0;
	    scanner2 = new Scanner(params);
		sc = scanner2.useDelimiter(",");
	    while (sc.hasNextInt()) {
	    	int val = sc.nextInt();
	    	if (numParms == 0) srcSheetNum = val;
	    	else if (numParms == 1) srcRowFrom = val; 
	    	else if (numParms == 2) srcRowTo = val; 
	    	else if (numParms == 3) srcColFrom = val; 
	    	else if (numParms == 4) srcColTo = val; 
	        numParms++;
	    }
		if (numParms != 5) {
			print_error("Error in argument "+(argNumber+1)+": Bad format of --srcArea. Expected 5 params: <sheet,RowFrom,RowTo,ColFrom,ColTo> but passed "+numParms);
			return;
		}
		argNumber++;
		
		// Parse --destCell=<sheet,RowTo,ColTo>
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.startsWith("--destcell=")) {
			print_error("Missing argument "+(argNumber+1)+":  --destCell=<sheet,RowTo,ColTo> => "+tmpArg);
			return;
		}	
		a = tmpArg.split("=");
		params = a[1];
		numParms=0;
	    sc = scanner2.useDelimiter(",");
	    while (sc.hasNextInt()) {
	    	int val = sc.nextInt();
	    	if (numParms == 0) destSheetNum = val;
	    	else if (numParms == 1) destRowTo = val; 
	    	else if (numParms == 2) destColTo = val; 
	        numParms++;
	    }
		if (numParms != 3) {
			print_error("Error in argument "+(argNumber+1)+": Bad format of --destCell. Expected 3 params: <sheet,RowTo,ColTo> but passed "+numParms);
			return;
		}
		argNumber++;
			
		// RegEx to check path syntax
		regex = Pattern.compile("^\"?("+REGEX_PATH+")\"?$",Pattern.CASE_INSENSITIVE);
		
		// Parse <srcFilePath1>
        srcPath1 = (argNumber<0 || argNumber>=args.length) ? "" : args[argNumber];
        m = regex.matcher(srcPath1);
        if (!m.find()) {
        	print_error("Empty or bad format for source file path  => " + srcPath1);
        	return;
        }
        srcPath1 = m.group(1);
        argNumber++;

		// Parse <srcFilePath2>
        srcPath2 = (argNumber<0 || argNumber>=args.length) ? "" : args[argNumber];
        m = regex.matcher(srcPath2);
        if (!m.find()) {
        	print_error("Empty or bad format for source file path  => " + srcPath2);
        	return;
        }
        srcPath2 = m.group(1);
        argNumber++;

		// Parse <destFilePath>
		destPath = (argNumber<0 || argNumber>=args.length) ? "" : args[argNumber];
        m = regex.matcher(destPath);
        if (!m.find()) {
        	print_error("Empty or bad format for destination file path  => " + destPath);
        	return;
        }
        destPath = m.group(1);
        argNumber++;
        
		/////////////////////////////////////////////////
		/////////////////////////////////////////////////
		Log.print (java.util.Arrays.toString(args));
		Log.print ("Comand line arguments OK with " + argNumber + " arguments.");
		Log.print ("   --srcParams =<" + srcSheetNum + "," + srcRowFrom + "," + srcRowTo + "," + srcColFrom + "," + srcColTo + ">");
		Log.print ("   --destParams=<" + destSheetNum + "," + destRowTo + "," + destColTo + ">");
		Log.print ("   srcPath = " + srcPath1);
		Log.print ("   srcPath = " + srcPath2);
		Log.print ("   destPath = " + destPath);
		/////////////////////////////////////////////////
		/////////////////////////////////////////////////
		
		File srcFile1 = new File(srcPath1);
		File srcFile2 = new File(srcPath2);
		File destFile = new File(destPath);

		Workbook srcBook = WorkbookFactory.create(srcFile1);
		Workbook destBook = WorkbookFactory.create(srcFile2);

		Sheet srcSheet = srcBook.getSheetAt(srcSheetNum);
		Sheet destSheet = destBook.getSheetAt(destSheetNum);

		System.out.println("copycells " + java.util.Arrays.toString(args));
		System.out.println ("");

	    System.out.println("Copying cell's area from " + srcFile1.getPath() + " to " + srcFile2.getPath() + " in " + destFile.getPath()); 		        
		copyArea (destBook, srcSheet, destSheet, srcRowFrom, srcRowTo, srcColFrom, srcColTo, destRowTo, destColTo, true);
	    try {
			FileOutputStream out = new FileOutputStream(destFile);
			destBook.write(out);
			out.close();
			destBook.close();
	    } catch (Exception e) {
	    	e.printStackTrace();
		}		
	    System.out.println("New Workbook created: " + destFile.getPath()); 		        
	    System.out.println();
	}
	
	/*
	 * Copy cells from scrSheet in the area denoted by: 
	 * upper-left corner (srcRowFrom, srcColFrom) and bottom-right corner (srcRowTo,srcColTo)
	 * To destSheet
	 * starting to copy at the upper-right corner (destRowTo, destColTo)
	 * If ...RowTo = -1 => ...RowTo = LastRow
	 * If ...ColTo = -1 => ...ColTo = lastCell
	 *  
	 */
	private static void copyArea (Workbook destWorkBook, Sheet srcSheet, Sheet destSheet, int srcRowFrom, int srcRowTo, int srcColFrom, int srcColTo, int destRowTo, int destColTo, boolean copyStyle) throws IOException {
		Map<Integer, CellStyle> styleMap = (copyStyle) ? new HashMap<Integer, CellStyle>() : null;
		
		if (srcRowTo < 0) srcRowTo = srcSheet.getLastRowNum();
		if (srcColTo < 0) {
			srcColTo = 0;
			Row r = srcSheet.getRow(srcRowFrom);
			if (r != null) 
				srcColTo = r.getLastCellNum();
		}
		if (destRowTo < 0) destRowTo = destSheet.getLastRowNum();
		if (destColTo < 0) {
			destColTo = 0;
			Row r = destSheet.getRow(destRowTo);
			if (r != null)
				destColTo = r.getLastCellNum();
		}

Log.print("Destination coordinates (row,col) = (" + destColTo + ","+destColTo+")");
		// Destination coordinates

		for (int i=srcRowFrom, count=0; i <= srcRowTo; i++, count++) {
Log.print("Copying row "+i);
			Row sRow = srcSheet.getRow(i);
			Row dRow = destSheet.getRow(destRowTo+count);
			if (sRow != null) {
				if (dRow == null) dRow = destSheet.createRow(destRowTo + count);
				WorkBookPlus.copyRowFromTo (destWorkBook, srcSheet, destSheet, i, srcColFrom, srcColTo, destRowTo + count, destColTo, styleMap);
			}
		}
	    for (int i = srcColFrom, count=0; i <= srcColTo; i++, count++)     
		      destSheet.setColumnWidth(destColTo+count, srcSheet.getColumnWidth(i));     
	} // End copyArea
	
	private static void print_error (String message) {
		System.out.println(CMD_USE);
		System.out.println(message);
		System.out.println();
	}
	
}
