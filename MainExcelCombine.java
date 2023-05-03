import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.jgm.general.Log;

public class MainExcelCombine {

	static String CMD_USE = "combine --operation=<SUM|DIF|MIN|MAX> --srcArea1=<Sheet,rowFrom,rowTo,colFrom,colTo> --srcArea2=<Sheet,rowFrom,colFrom> <srcFilePath1> <srcFilePath2> <destFilePath>\n" +
							"\n";
	public enum Operations {SUM, DIF, MIN, MAX};

	final static String REGEX_PATH = "([a-zA-Z]:)?([\\\\]{0,2}[a-zA-Z0-9 _.-]+)+[\\\\]{0,2}";
	private static Scanner scanner;

	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception {

		Log.setup(false,"");
		
		String tmpArg = "";
		String srcPath1 = "";
		String srcPath2 = "";
		String destPath = "";
		int src1SheetNum = 0;
		int src1RowFrom = 0;
		int src1RowTo=0;
		int src1ColFrom = 0;
		int src1ColTo=0;
		int src2SheetNum = 0;
		int src2RowFrom = 0;
		int src2ColFrom = 0;
		Operations op;

		String a[]; // Temp array for splitting strings
		String params; // Comma separated string of sheet,row,col...
		int numParms;
		Scanner sc; 
		Pattern regex;
		Matcher m;
		int argNumber = 0;
		
		int nArgs = args.length;
		if (nArgs < 6 || nArgs > 6) {
			print_error("Bad number of arguments.");
			return;
		}
		
		// Parse --operation=<SUM|DIF|MIN|MAX>
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.startsWith("--operation=")) {
			print_error("Missing argument "+(argNumber+1)+":  --operation=<SUM|DIF|MIN|MAX> => "+tmpArg);
			return;
		}	
		a = tmpArg.split("=");
		params = a[1];
		if (params.equals("min")) op = Operations.MIN;
		else if (params.equals("max")) op = Operations.MAX;
		else if (params.equals("sum")) op = Operations.SUM;
		else if (params.equals("dif")) op = Operations.DIF;
		else {
			print_error("Error in argument "+(argNumber+1)+": Bad format of --operation. Expected <SUM|DIF|MIN|MAX> and got "+params);
			return;
		}
		argNumber++;
		
		// Parse --srcArea1=<sheet,RowFrom,RowTo,ColFrom,ColTo>
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.startsWith("--srcarea1=")) {
			print_error("Missing argument "+(argNumber+1)+":  --srcArea1=<sheet,RowFrom,RowTo,ColFrom,ColTo> => "+tmpArg);
			return;
		}	
		a = tmpArg.split("=");
		params = a[1];
		numParms=0;
	    scanner = new Scanner(params);
		sc = scanner.useDelimiter(",");
	    while (sc.hasNextInt()) {
	    	int val = sc.nextInt();
	    	if (numParms == 0) src1SheetNum = val;
	    	else if (numParms == 1) src1RowFrom = val; 
	    	else if (numParms == 2) src1RowTo = val; 
	    	else if (numParms == 3) src1ColFrom = val; 
	    	else if (numParms == 4) src1ColTo = val; 
	        numParms++;
	    }
		if (numParms != 5) {
			print_error("Error in argument "+(argNumber+1)+": Bad format of --srcArea. Expected 5 params: <sheet,RowFrom,RowTo,ColFrom,ColTo> but passed "+numParms);
			return;
		}
		argNumber++;
		
		// Parse --srcArea2=<sheet,RowFrom,ColFrom>
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.startsWith("--srcarea2=")) {
			print_error("Missing argument "+(argNumber+1)+":  --srcArea2=<sheet,RowFrom,ColFrom> => "+tmpArg);
			return;
		}	
		a = tmpArg.split("=");
		params = a[1];
		numParms=0;
	    sc = scanner.useDelimiter(",");
	    while (sc.hasNextInt()) {
	    	int val = sc.nextInt();
	    	if (numParms == 0) src2SheetNum = val;
	    	else if (numParms == 1) src2RowFrom = val; 
	    	else if (numParms == 2) src2ColFrom = val; 
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
		Log.print (java.util.Arrays.toString(args));
		Log.print ("Comand line arguments OK with " + argNumber + " arguments.");
		Log.print ("   --srcArea1 =<" + src1SheetNum + "," + src1RowFrom + "," + src1RowTo + "," + src1ColFrom + "," + src1ColTo + ">");
		Log.print ("   --srcArea2=<" + src2SheetNum + "," + src2RowFrom + "," + src2ColFrom + ">");
		Log.print ("   srcPath = " + srcPath1);
		Log.print ("   srcPath = " + srcPath2);
		Log.print ("   destPath = " + destPath);
		/////////////////////////////////////////////////
		
		// Now get DestinationFilePath SourceFilePath FilterFilePath args	
		File srcFile1 = new File(srcPath1);
		File srcFile2 = new File(srcPath2);
		File destFile = new File(destPath);
		
		Workbook src1Book = WorkbookFactory.create (srcFile1);
		Workbook src2Book = WorkbookFactory.create (srcFile2);
		Sheet src1Sheet = src1Book.getSheetAt(src1SheetNum);
		Sheet src2Sheet = src2Book.getSheetAt(src2SheetNum);

		System.out.println("combine " + java.util.Arrays.toString(args));
		System.out.println ("");

		combine (src1Sheet, src2Sheet, src1RowFrom, src1RowTo, src1ColFrom, src1ColTo, src2RowFrom, src2ColFrom, op);
		
		try {
			Workbook destBook = src1Book;
			OutputStream out = new FileOutputStream(destFile);
			destBook.write(out);
			out.close();
			destBook.close();
		} catch( IOException e ) {
            e.printStackTrace();
		}
	    System.out.println("New Workbook created: " + destFile.getPath()); 		        
	    System.out.println();
	}
	
	private static void combine (Sheet src1Sheet, Sheet src2Sheet, int src1RowFrom, int src1RowTo, int src1ColFrom, int src1ColTo, int src2RowFrom, int src2ColFrom, Operations op) throws IOException {
		if (src1RowTo < 0) src1RowTo = src1Sheet.getLastRowNum();
		if (src1ColTo < 0) {
			src1ColTo = 0;
			Row r = src1Sheet.getRow(src1RowFrom);
			if (r != null) 
				src1ColTo = r.getLastCellNum();
		}
		if (src2RowFrom < 0) src2RowFrom = src2Sheet.getLastRowNum();
		if (src2ColFrom < 0) {
			src2ColFrom = 0;
			Row r = src2Sheet.getRow(src2RowFrom);
			if (r != null)
				src2ColFrom = r.getLastCellNum();
		}

		for (int i=src1RowFrom, countRows=0; i <= src1RowTo; i++, countRows++) {
			double newVal = 0, s1CellVal, s2CellVal;
			int s1Type=0,s2Type=0;
			Row s1Row = src1Sheet.getRow(i);
			Row s2Row = src2Sheet.getRow(src2RowFrom+countRows);
			
			Cell s1Cell = null, s2Cell = null, dCell = null;
			for (int j=src1ColFrom, countCols=0; j <= src1ColTo; j++, countCols++) {
				s1CellVal = 0;
				if (s1Row != null) {
					s1Cell = s1Row.getCell(j);
					if (s1Cell != null ) {
						s1Type = s1Cell.getCellType();
						if (s1Type == Cell.CELL_TYPE_NUMERIC)
							s1CellVal = s1Cell.getNumericCellValue();
					}	
				} 
				
				s2CellVal = 0;
				if (s2Row != null) {
					s2Cell = s2Row.getCell(src2ColFrom+countCols);
					if (s2Cell != null ) {
						s2Type = s2Cell.getCellType();
						if (s2Type == Cell.CELL_TYPE_NUMERIC)
							s2CellVal = s2Cell.getNumericCellValue();
					}	
				}
				switch (op) {
				case SUM:
					newVal = s1CellVal + s2CellVal;
					break;
				case DIF:
					newVal = s1CellVal - s2CellVal;
					break;
				case MIN:
					newVal = (s1CellVal < s2CellVal) ? s1CellVal : s2CellVal;
					break;
				case MAX:
					newVal = (s1CellVal > s2CellVal) ? s1CellVal : s2CellVal;
					break;
				} // EndSwitch
				dCell = s1Cell;
				if (dCell != null && s1Type == s2Type) { 
					dCell.setCellValue(newVal);
Log.print ("Setting value for dest ("+i+","+j+") = "+newVal);
				}
				else {
Log.print ("Cell types from src1 ("+i+","+j+") and src2 ("+((int)src2RowFrom+countRows)+","+((int)src2ColFrom+countCols)+") are not the same !!!");
Log.print ("Celltypes Src1: " + s1Type + " - Src2: " + s2Type);
				}
			} //EndFor j
		} // EndFor i
	}
	
	
	private static void print_error (String message) {
		System.out.println(CMD_USE);
		System.out.println(message);
		System.out.println();
	}
}
