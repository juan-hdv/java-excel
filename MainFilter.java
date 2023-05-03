import java.io.File;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.indracompany.general.Log;

import excelfilter.ExcelFilter;

public class MainFilter {

	enum FilterType {NIL, INFILTER, NOTINFILTER, FROMPHRASES};
	static String CMD_USE = 
			"USE 1:\n" +
			"  filter -d\n" +
			" 		 --InFilter|--NotInFilter\n" +
			"		 --srcParams=<sheetNum,rowNum,colNum>\n" +
			"		 --filterParams=<sheetNum,rowNum,colNum>\n" +
			"		 SourceFilePath\n" +
			"		 FilterFilePath\n" +
			"		 DestinationFilePath\n" +
			"* Delete all rows from <SourceFilePath> where\n" +
			"  the contents of col <colNum> of sheet <sheetNum>, starting at row <rowNum>,\n" +
			"  is present (<InFilter>) or is not (<NotInfilter>) in <FilterFilePath>\n" +
			"  at Sheet <SheetNum> and col <colNum>, starting at row <rowNum>.\n" +
			"  The new file is saved to <DestinationFilePath>.\n\n" +
			" USE 2:\n" +
			"   filter -d\n" +
			"		  --Phrases=<string1,string2,string3...>\n" +
			"         --srcParams=<sheetNum,rowNum,colNum>\n" +
			"         SourceFilePath\n" +
			"		  DestinationFilePath \n" +
			"* Delete all rows from <SourceFilePath> where\n" +
			"  the contents of col <colNum> of sheet <sheetNum>, starting at row <rowNum>,\n" +
			"  equals any of the phrases <string1..stringN>\n" +
			"  The new file is saved to <DestinationFilePath>.\n";
	
	final static String REGEX_PATH = "([a-zA-Z]:)?([\\\\]{0,2}[a-zA-Z0-9 _.-]+)+[\\\\]{0,2}";

	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception {

		Log.setup(false,"");
		
		String tmpArg = "";
		String destPath="", srcPath="", filterPath="";
		String[] phrases = null;
		FilterType filterType;
		int srcSheetNum=-1,srcRowNum=-1,srcColNum=-1;
		int filterSheetNum=-1,filterRowNum=-1,filterColNum=-1;
		String a[]; // Temp array for splitting strings
		String params; // Comma separated string of sheet,row,col
		int numParms;
		Scanner sc; 
		
		Pattern regex;
		Matcher m;
		int argNumber = 0;
		
		int nArgs = args.length;
		if (nArgs < 5 || nArgs > 7) {
			print_error("Bad number of arguments: " + nArgs);
			return;
		}
		
		// Parse -d
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.equals("-d")) {
			print_error("Missing argument "+(argNumber+1)+": -d flag.");
			return;
		}
		argNumber++;
		
		// Parse --InFilter|--NotInFilter|--Phrases
		tmpArg = args[argNumber].toLowerCase();
		filterType = FilterType.NIL;
		if (tmpArg.startsWith("--infilter")) {
			filterType = FilterType.INFILTER; 
		} else if (tmpArg.startsWith("--notinfilter")) {
			filterType = FilterType.NOTINFILTER;
		} else if (tmpArg.startsWith("--phrases")) {
			filterType = FilterType.FROMPHRASES;
    		a = tmpArg.split("=");
    		phrases = a[1].split(",");
    		if (phrases.length == 0) {
	        	print_error("Argument "+(argNumber+1)+"(Phrases) bad format. Must be Phrases=<String1...StringN>");
	        	return;
    		}
		} else {
			print_error("Missing argument "+(argNumber+1)+":  --InFilter|--NotInFilter|--Phrases=<Phr1,..Phrn> flag.");
			return;
		}
		argNumber++;
		
		// Now parse --srcParams=<sheetNum,rowNum,colNum>
		tmpArg = args[argNumber].toLowerCase();
		if (!tmpArg.startsWith("--srcparams=")) {
			print_error("Missing argument "+(argNumber+1)+":  --srcParams=<sheetNum,rowNum,colNum> => "+tmpArg);
			return;
		}	
		a = tmpArg.split("=");
		params = a[1];
		numParms=0;
	    sc = new Scanner(params).useDelimiter(",");
	    while (sc.hasNextInt()) {
	    	int val = sc.nextInt();
	    	if (numParms == 0) srcSheetNum = val;
	    	else if (numParms == 1) srcRowNum = val; 
	    	else if (numParms == 2) srcColNum = val; 
	        numParms++;
	    }
		if (numParms != 3) {
			print_error("Error in argument "+(argNumber+1)+": Bad format of --srcParams. Expected 3 params: <sheetNum,rowNum,colNum> but passed "+numParms);
			return;
		}
		argNumber++;
		
		// If filterType is --InFilter|--NotInFilter then parse --filterParams=<sheetNum,rowNum,colNum>
		if (filterType == FilterType.INFILTER || filterType == FilterType.NOTINFILTER) {
			tmpArg = args[argNumber].toLowerCase();
			if (!tmpArg.startsWith("--filterparams=")) {
				print_error("Missing argument "+(argNumber+1)+":  --filterParams=<sheetNum,rowNum,colNum>");
				return;
			}	
			a = tmpArg.split("=");
			params = a[1];
			numParms=0;
		    sc = new Scanner(params).useDelimiter(",");
		    while (sc.hasNextInt()) {
		    	int val = sc.nextInt();
		    	if (numParms == 0) filterSheetNum = val;
		    	else if (numParms == 1) filterRowNum = val; 
		    	else if (numParms == 2) filterColNum = val; 
		        numParms++;
		    }
			if (numParms != 3) {
				print_error("Error in argument "+(argNumber+1)+": Bad format of --filterParams. Expected 3 params: <sheetNum,rowNum,colNum> but passed "+numParms);
				return;
			}
			argNumber++;
		}
		// PRE: Next parse arguments are:  SourceFilePath [FilterFilePath ]DestinationFilePath
			
		// RegEx to check path syntax
		regex = Pattern.compile("^\"?("+REGEX_PATH+")\"?$",Pattern.CASE_INSENSITIVE);
		
		// Source (input) file 
        srcPath = (argNumber<0 || argNumber>=args.length) ? "" : args[argNumber];
        m = regex.matcher(srcPath);
        if (!m.find()) {
        	print_error("Empty or bad format for source file path  => " + srcPath);
        	return;
        }
        srcPath = m.group(1);
        argNumber++;
        
        // Filter file; Only if FilterType == FROMFILE
        if (filterType == FilterType.INFILTER || filterType == FilterType.NOTINFILTER) { 
	        filterPath = (argNumber<0 || argNumber>=args.length) ? "" : args[argNumber];
	        m = regex.matcher(filterPath);
	        if (!m.find()) {
	        	print_error("Empty or bad format for filter file path  => " + filterPath);
	        	return;
	        }
			filterPath = m.group(1);
	        argNumber++;
        }
		
		// Destination (output) file
		destPath = (argNumber<0 || argNumber>=args.length) ? "" : args[argNumber];
        m = regex.matcher(destPath);
        if (!m.find()) {
        	print_error("Empty or bad format for destination file path  => " + destPath);
        	return;
        }
        destPath = m.group(1);
        
/////////////////////////////////////////////////
/////////////////////////////////////////////////
Log.print (java.util.Arrays.toString(args));
Log.print ("Comand line arguments OK with " + argNumber + " arguments.");
switch (filterType) {
	case INFILTER:	  Log.print ("   FilterType = InFilter");
		break;
	case NOTINFILTER: Log.print ("   FilterType = NotInFilter");
		break;
	case FROMPHRASES: Log.print ("   FilterType = FromPhraseS");
		break;
}
Log.print ("   --srcParams.SheetNum = " + srcSheetNum);
Log.print ("   --srcParams.RowNum = " + srcRowNum);
Log.print ("   --srcParams.ColNum = " + srcColNum);
switch (filterType) {
	case INFILTER:
	case NOTINFILTER: 
		Log.print ("   --filterParams.SheetNum = " + filterSheetNum);
		Log.print ("   --filterParams.RowNum = " + filterRowNum);
		Log.print ("   --filterParams.ColNum = " + filterColNum);
		break;
}
Log.print ("   srcPath = " + srcPath);
Log.print ("   destPath = " + destPath);
switch (filterType) {
	case INFILTER:
	case NOTINFILTER:
		Log.print ("   filterPath = " + filterPath);
		break;
	case FROMPHRASES:
		Log.print ("   Phrases = " + StringUtils.join (phrases,","));
		break;
}
/////////////////////////////////////////////////
/////////////////////////////////////////////////

		System.out.println("filter " + java.util.Arrays.toString(args));

  		ExcelFilter ef = new ExcelFilter();
		
		File srcFile = new File(srcPath);
		File destFile = new File(destPath);
		
		int numRowsOrg = WorkbookFactory.create(srcFile).getSheetAt(srcSheetNum).getLastRowNum()+1;
		int resultingRows = 0;
		boolean in = false;
		switch (filterType) {
			case INFILTER: 
				in = true;
			case NOTINFILTER:
				File filterFile = new File(filterPath);
				if (in) {
					resultingRows = ef.DeleteRowsInFilter(srcFile, destFile, filterFile, srcSheetNum, srcRowNum, srcColNum, filterSheetNum, filterRowNum, filterColNum);
				} else {
					resultingRows = ef.DeleteRowsNotInFilter(srcFile, destFile, filterFile, srcSheetNum, srcRowNum, srcColNum, filterSheetNum, filterRowNum, filterColNum);
				}
				break;
			case FROMPHRASES:
				resultingRows = ef.DeleteRowsContainingPhrases (srcFile, destFile, phrases, srcSheetNum, srcRowNum, srcColNum);
				break;
		}	
	    System.out.println("Deleted rows:   " + (numRowsOrg-resultingRows)); 		        
	    System.out.println("Resulting rows: " + resultingRows); 		        
	    System.out.println("New Workbook created: " + destFile.getPath()); 		        
	    System.out.println();
	}
	
	private static void print_error (String message) {
		System.out.println(CMD_USE);
		System.out.println(message);
		System.out.println();
	}
	
}
