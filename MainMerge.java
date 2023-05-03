import java.io.File;
import java.io.FileInputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import excelmerge.ExcelMerge;

import com.indracompany.general.Log;


public class MainMerge {
	
	final static String CMD_USE = "USE: merge out=path sheet1=name1,path11,path12 sheet2=name2,path21,path22...";
	
	/**
	 * @param args
	 * @return 
	 */
	public static void main(String[] args) throws Exception {
		
		String[] sheetNameList;
		FileInputStream[][] fileList; 
		int nArgs = args.length;
		
		String foutname = "";
		if (nArgs < 2) {
			print_error("Bad number of arguments");
			return;
		}
		String path = "([a-zA-Z]:)?([\\\\]{0,2}[a-zA-Z0-9 _.-]+)+[\\\\]{0,2}";
		
		// 1. Validate out file arg syntax
		Pattern regex = Pattern.compile("^out=\"?("+path+")\"?$",Pattern.CASE_INSENSITIVE);
        Matcher m = regex.matcher(args[0]);
        if (!m.find()) {
        	print_error("Bad out= format => " + args[0]);
        	return;
        }
        foutname = m.group(1);

		sheetNameList = new String[nArgs];
		fileList = new FileInputStream[nArgs][];
		// 2. Validate sheet arg syntax
        // sheetK=name,path
        for (int k=1; k<nArgs;k++) {
			String[] parts = args[k].split("=");
			if (parts.length != 2) {
				print_error("Wrong sheet= format. Spliting by '=' results in more than 2 parts");
				return;
			}
			String sheetToken = parts[0];
			String paramsToken = parts[1];
			// Check if 'sheet[0-9]+' is present (sheetToken)
			regex = Pattern.compile("^sheet[0-9]{1,2}$",Pattern.CASE_INSENSITIVE);
			m = regex.matcher(sheetToken);
	        if (!m.find()) {
	        	print_error("Wrong sheet= format. 'Sheet[0-9]+' must be present");
	        	return;
	        }
			// Split paramsToken and Check if sheet 'name' and 'paths' are present and are valid
			parts = paramsToken.split(",");
			if (parts.length < 2) {
				print_error("Wrong sheet= format. Spliting by ',' results in less than 2 parts");
				return;
			}

			String sheetNameToken = parts[0];
	        // Check if first param is a sheet name
			regex = Pattern.compile("\"?([a-zA-Z0-9 _-]+)\"?",Pattern.CASE_INSENSITIVE);
			m = regex.matcher(sheetNameToken);
	        if (!m.find()) {
	        	print_error("Wrong sheet"+k+"= format. Sheet name missing");
	        	return;
	        }
			sheetNameList[k-1] =  m.group(1); // Sheet name
	        
	        // Check if the rest of params are paths
	        int numPaths = parts.length-1;
			fileList[k-1] = new FileInputStream[numPaths];
			for (int j=1; j<=numPaths; j++) {
				String thePath = parts[j];  
				// Check if is a valid path
				regex = Pattern.compile("^\"?("+path+")\"?$",Pattern.CASE_INSENSITIVE);
				m = regex.matcher(thePath);
				if (!m.find()) {
					print_error("Wrong sheet"+k+"= format. Bad path format => " + thePath);
					return;
				}
				thePath = m.group (1);
				fileList[k-1][j-1] = new FileInputStream(thePath); // Path
			}
        } // End for
        
		System.out.println("merge " + java.util.Arrays.toString(args));
		System.out.println ("");
		
        Log.setup (false,"");
		File fout = new File(foutname);
		ExcelMerge em = new ExcelMerge();
		em.mergeExcelFiles(fout, fileList, sheetNameList);
	}
	
	private static void print_error (String message) {
		System.out.println(CMD_USE);
		System.out.println(message);
		System.out.println();
	}

}
