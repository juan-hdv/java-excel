import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.jgm.general.Log;

public class ExcelCompare {

	static String CMD_USE = "compare [-exact] --srcCol1=<Sheet,rowFrom,rowTo,col> --srcCol2=<Sheet,rowFrom,col> <srcFilePath1> <srcFilePath2>\n" +
			"-exact compares row by row (1 to 1 comparison)\n" +
			"when -exact is not present, each element from src1 is searched in the col of src2 (1 to N comparison)\n" +
			"--srcCol1=<Sheet,rowFrom,rowTo,col> Parameters of the Source 1 sheet" +
			"--srcCol2=<Sheet,rowFrom,col> Parameters of the Source 2 sheet";

	final static String REGEX_PATH = "([a-zA-Z]:)?([\\\\]{0,2}[a-zA-Z0-9 _.-]+)+[\\\\]{0,2}";
	
	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception {

		// Log.setup(true,"log.log");
		Log.setup(false,"");
		
		String tmpArg = "";
		String srcPath1 = "";
		String srcPath2 = "";
		int src1SheetNum = 0;
		int src1RowFrom = 0;
		int src1RowTo=0;
		int src1Col = 0;
		int src2SheetNum = 0;
		int src2RowFrom = 0;
		int src2Col = 0;
		boolean exact = false;
		boolean caseSensitive = false;

		String a[]; // Temp array for splitting strings
		String params; // Comma separated string of sheet,row,col...
		int numParms;
		Scanner sc; 
		Pattern regex;
		Matcher m;
		int argNumber = 0;
		
		int nArgs = args.length;
		if (nArgs < 4 || nArgs > 5) {
			print_error("Bad number of arguments.");
			return;
		}
		
		// Parse [-exact]
		tmpArg = args[argNumber].toLowerCase();
		if (tmpArg.equals("-exact")) {
			exact = true;
			argNumber++;
		}	
		
		// Parse --srcCol1=<Sheet,rowFrom,rowTo,col>
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.startsWith("--srccol1=")) {
			print_error("Missing argument "+(argNumber+1)+":  --srcCol1=<Sheet,rowFrom,rowTo,col> => "+tmpArg);
			return;
		}	
		a = tmpArg.split("=");
		params = a[1];
		numParms=0;
	    sc = new Scanner(params).useDelimiter(",");
	    while (sc.hasNextInt()) {
	    	int val = sc.nextInt();
	    	if (numParms == 0) src1SheetNum = val;
	    	else if (numParms == 1) src1RowFrom = val; 
	    	else if (numParms == 2) src1RowTo = val; 
	    	else if (numParms == 3) src1Col = val; 
	        numParms++;
	    }
		if (numParms != 4) {
			print_error("Error in argument "+(argNumber+1)+": Bad format of --srcCol1. Expected 4 params: <Sheet,rowFrom,rowTo,col> but passed "+numParms);
			return;
		}
		argNumber++;
		
		// Parse --srcCol2=<Sheet,rowFrom,col>
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.startsWith("--srccol2=")) {
			print_error("Missing argument "+(argNumber+1)+":  --srcCol2=<Sheet,rowFrom,col> => "+tmpArg);
			return;
		}	
		a = tmpArg.split("=");
		params = a[1];
		numParms=0;
	    sc = new Scanner(params).useDelimiter(",");
	    while (sc.hasNextInt()) {
	    	int val = sc.nextInt();
	    	if (numParms == 0) src2SheetNum = val;
	    	else if (numParms == 1) src2RowFrom = val; 
	    	else if (numParms == 2) src2Col = val; 
	        numParms++;
	    }
		if (numParms != 3) {
			print_error("Error in argument "+(argNumber+1)+": Bad format of --srcCol2. Expected 3 params: <Sheet,rowFrom,col> but passed "+numParms);
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

		/////////////////////////////////////////////////
		Log.print (java.util.Arrays.toString(args));
		Log.print ("Comand line arguments OK with " + argNumber + " arguments.");
		Log.print ("   --srcCol1 =<" + src1SheetNum + "," + src1RowFrom + "," + src1RowTo + "," + src1Col + ">");
		Log.print ("   --srcCol2=<" + src2SheetNum + "," + src2RowFrom + "," + src2Col + ">");
		Log.print ("   src1Path = " + srcPath1);
		Log.print ("   src2Path = " + srcPath2);
		/////////////////////////////////////////////////
		
		// Now get DestinationFilePath SourceFilePath FilterFilePath args	
		File src1File = new File(srcPath1);
		File src2File = new File(srcPath2);
		
		Sheet src1Sheet = WorkbookFactory.create (src1File).getSheetAt(src1SheetNum);
		Sheet src2Sheet = WorkbookFactory.create (src2File).getSheetAt(src2SheetNum);
		
		System.out.println("compare " + java.util.Arrays.toString(args));
		System.out.println ("");
		if (exact)
			compare_1x1 (src1Sheet, src2Sheet, src1RowFrom, src1RowTo, src1Col, src2RowFrom, src2Col, caseSensitive);
		else 
			compare_NxN (src1Sheet, src2Sheet, src1RowFrom, src1RowTo, src1Col, src2RowFrom, src2Col, caseSensitive);
	}
	
	/*
	 * True if:
	 *     - There is not row from src1 different to the corresponding src2 row
	 *     - The number of existing (not null) rows are the same in each book-sheet (src1 and src2)     
	 */
	private static boolean compare_1x1 (Sheet src1Sheet, Sheet src2Sheet, int src1RowFrom, int src1RowTo, int src1Col, int src2RowFrom, int src2Col, boolean sensitive) throws IOException {
		int src1NumRows, src2NumRows;
		int src2RowTo;

		int RowTo = -1;
		if (src1RowTo == -1) {
			src1RowTo = src1Sheet.getLastRowNum();
			src2RowTo = src2Sheet.getLastRowNum();
		} else
			src2RowTo = src2RowFrom + (src1RowTo - src1RowFrom);
     	src1NumRows = src1RowTo - src1RowFrom + 1;
		src2NumRows = src2RowTo - src2RowFrom + 1;
		RowTo = src1RowTo > src2RowTo ? src1RowTo : src2RowTo;

		String val1=null, val2=null;
		int differences = 0, similarities = 0;
		int count = 0;
		for (int i=src1RowFrom; i <= RowTo; i++, count++) {
			int j = src2RowFrom + count;
			val1 = getStringVal(src1Sheet,i, src1Col);
			if (val1 == null || val1.equals("0.0")) val1 = ""; // Para valores numéricos, solución temporal
			val2 = getStringVal(src2Sheet,j, src2Col);
			if (val2 == null || val2.equals("0.0")) val2 = "";
			if (sensitive ? val1.equals(val2) : val1.equalsIgnoreCase(val2))
				similarities++;
			else {
				differences++;
				System.out.println ("Src1("+i+","+src1Col+")["+val1+"] != Src2("+j+","+src2Col+")["+val2+"]");
				System.out.println ("     Src1("+i+",1)["+val1+"] >> [" + getStringVal(src1Sheet,i,1) + "]");
				System.out.println ("     Src2("+j+",1)["+val2+"] >> [" + getStringVal(src2Sheet,j,1) + "]");
//				System.out.println ("Src1("+i+","+src1Col+")["+getStringVal(src1Sheet,i, 1)+"->"+val1+"] != Src2("+j+","+src2Col+")["+getStringVal(src2Sheet,j, 1)+"->"+val2+"]");
				
			}
		}
		boolean result = (differences==0) && (similarities == count) && (src1NumRows == src2NumRows); 

		if (differences!=0) System.out.println ("");
		System.out.println ("Summary Comparing 1x1 - Rows must be in the same order");
		System.out.println ("=================");
		System.out.println ("Num rows looped: "+count);
		System.out.println ("Differences in # of rows: "+Math.abs(src1NumRows-src2NumRows));
		System.out.println ("Differences (values): "+differences);
		System.out.println ("Similarities (values): "+similarities);
		System.out.println ("SRC1 == SRC2 ? "+ (result?"***TRUE***":"***FALSE***"));
		return result;
	}

	/*
	 * True if:
	 *     - There is each row in src1 is in scr2 and viceversa
	 *     - The number of existing (not null) rows are the same in each book-sheet (src1 and src2)     
	 */
	private static boolean compare_NxN (Sheet src1Sheet, Sheet src2Sheet, int src1RowFrom, int src1RowTo, int src1Col, int src2RowFrom, int src2Col, boolean sensitive) throws IOException {
		int src1NumRows, src2NumRows;
		int src2RowTo;
		Map<String,Integer> src1Map = new HashMap<String, Integer>(); 
		Map<String,Integer> src2Map = new HashMap<String, Integer>(); 
		
		int RowTo = -1;
		if (src1RowTo == -1) {
			src1RowTo = src1Sheet.getLastRowNum();
			src2RowTo = src2Sheet.getLastRowNum();
		} else
			src2RowTo = src2RowFrom + (src1RowTo - src1RowFrom);
     	src1NumRows = src1RowTo - src1RowFrom + 1;
		src2NumRows = src2RowTo - src2RowFrom + 1;
		RowTo = src1RowTo > src2RowTo ? src1RowTo : src2RowTo;
Log.print ("src1NumRows=" + src1NumRows + " src2NumRows="+src2NumRows);
Log.print ("src1RowTo=" + src1RowTo + " src2RowTo="+src2RowTo);
		
		// Map each element from Src1 and Src2 columns: +1 if from scr1, -1 if from src2
		for (int i=src1RowFrom, j=src2RowFrom; i <= RowTo; i++, j++) {
			if (i <= src1RowTo) {
				String key1 = getStringVal(src1Sheet,i, src1Col);
				if (key1 == null) key1 ="null";
				Integer val1 = src1Map.get(key1);
				src1Map.put (key1, (val1 == null ? 1 : val1+1));
			}
			if (j <= src2RowTo) {
				String key2 = getStringVal(src2Sheet,j, src2Col);
				if (key2 == null) key2 ="null";
				Integer val2 = src2Map.get(key2);
				src2Map.put (key2, (val2 == null ? 1 : val2+1));
			}
		} // end for

Log.print ("");
Log.print ("Lista de personas Src1: " + src1Map.toString());
Log.print ("");
Log.print ("Lista de personas Src2: " + src2Map.toString());
Log.print ("");

		// Compares the 2 Maps, to find similarities and differences 
		int differences = 0, similarities = 0;
		Set<String> union = new HashSet<String>(src1Map.keySet());
		union.addAll(src2Map.keySet());
Log.print ("Elements SET size = "+union.size());		
		for (String myKey : union) {
Log.print ("For each entry in Set. Entry = "+myKey);
			Integer	val1 = src1Map.get(myKey); // Integer accepts NULL , int do not
			val1 = val1==null?0:val1;
			Integer val2 = src2Map.get(myKey);
			val2 = val2==null?0:val2;
Log.print ("src1Map.get("+myKey+") = " + val1);			
Log.print ("src2Map.get("+myKey+") = " + val2);			
			if (val1 == val2) {
				similarities+= val1;
			} else {
				int delta = Math.abs(val1-val2);
				differences += delta;
				similarities += Math.min(val1,val2);
				System.out.println ("["+myKey+"]" + " is missing from " + (val1 > val2 ? "src2 ":"src1 ") + delta + " times.");
			}
		}
		int numRowsLooped = RowTo- src1RowFrom + 1;
		boolean result = (differences==0) && (similarities == numRowsLooped) && (src1NumRows == src2NumRows);

		if (differences!=0) System.out.println ("");
		System.out.println ("Summary Comparing NxN - Rows don't have to be in the same order");
		System.out.println ("=================");
		System.out.println ("Num rows looped: "+numRowsLooped);
		System.out.println ("Differences in # of rows: "+Math.abs(src1NumRows - src2NumRows));
		System.out.println ("Differences  : "+differences);
		System.out.println ("Similarities : "+similarities);
		System.out.println ("SRC1 == SRC2 ? "+ (result?"***TRUE***":"***FALSE***"));
		return result;
	}

	private static String getStringVal (Sheet sheet, int rowNum, int colNum) {
		String val = null;
		Row r = sheet.getRow(rowNum);
		if (r == null) return null;
		Cell c = r.getCell(colNum);
		if (c == null) return null;
		
		int type = c.getCellType();
		if (type == Cell.CELL_TYPE_FORMULA)
			type = c.getCachedFormulaResultType();
		switch (type) {
			case Cell.CELL_TYPE_BOOLEAN:
				val = (c.getBooleanCellValue()?"true":"false");
				break;
			case Cell.CELL_TYPE_STRING:
				val = c.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				val = Double.toString(c.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
			case Cell.CELL_TYPE_ERROR:
			case Cell.CELL_TYPE_BLANK:
				val = "";
				break;
			default:
				val = null;
		}
		return (val==null?null:val.trim());
	}
	
	private static void print_error (String message) {
		System.out.println(CMD_USE);
		System.out.println(message);
		System.out.println();
	}
}
