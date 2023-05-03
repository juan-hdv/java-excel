import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author 	jgmejia
 * @date 	2014/11
 *
 */
public class Programa {

	final static boolean DEBUG = true;
	static String templatePath = "Template/";
	static String dataPath = "data/";
	static String filenameOut = dataPath + "Resumen_Control_Produccion_%s.xls";
	//static String filenameOut = dataPath + "temp %s.xls";
	
	static String filenameTemplate = templatePath + "Resumen_Control_Produccion_template.xls";
	static Map<String, Integer> riesgo = new TreeMap<String, Integer>();
	static Map<String, Integer> impactoEntrega = new TreeMap<String, Integer>();
	static String[][] resultSheet;

	final static int FILE_ROW_FIRST = 7; 
	// Columnas de control
	final static int FILE_COL_SUBTAREA = 3; // Key subtarea
	
	final static int MAT_NUM_COLS = 5;
	final static int[] ColsArray = new int[MAT_NUM_COLS]; 
	final static int COL_RIESGO = 0;
	final static int COL_IMPACTO = 1;
	final static int COL_TIPO_CAUSA = 2;
	final static int COL_ACCIONES_MITIGACION = 3;
	final static int COL_OBSERVACIONES = 4;

	
	private static void initializeArrays () {
		ColsArray[COL_RIESGO] = 21;
		ColsArray[COL_IMPACTO] = 22;
		ColsArray[COL_TIPO_CAUSA] = 23;
		ColsArray[COL_ACCIONES_MITIGACION] = 25;
		ColsArray[COL_OBSERVACIONES] = 26;
	} // End initializeArrays
	
	private static void initializeHashes () {
		riesgo.put("No hay Riesgo de Avance", 0);
		riesgo.put("Problema productividad", 0);
		riesgo.put("Retraso en Planificacion", 0);
		riesgo.put("Sin Planificación", 0);		
		
		impactoEntrega.put("SI", 0);
		impactoEntrega.put("NO", 0);
	} // End initializeHashes

	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception {
		String[][] projectSheet;
		String filenameIn;
		String[] abstractxSheet;
		
		initializeArrays();
		
		Workbook template_wb = WorkbookFactory.create(new File(filenameTemplate));
		Sheet template_ws = template_wb.getSheetAt(0);
		
		ArrayList<String> filenames = getFilenames4Folder (dataPath, "Control*.xls");
		ArrayList<String> fileDate = new ArrayList<String>();
		Iterator<String> iterator = filenames.iterator();
		String date = "";
		int numProys = 0;
		int col = 1;
		// Recorre todos los archivos del directorio
		while (iterator.hasNext()) {
			filenameIn = dataPath  + iterator.next();
			log ("Procesando... " + filenameIn);		
			// Get the project name from filename. \\p{L} es cualquier letra UNICODE en cualquier lenguaje
			Pattern regExp = Pattern.compile("Control de Produccion\\s([\\p{L}_ -]+)\\s([0-9]{8}).xls$");
			Matcher m = regExp.matcher(filenameIn);
			String project = "";
			if (m.find()) {
				project = m.group(1);
				date = m.group(2);
				if (project == null || date == null || project.length() == 0 || date.length() == 0) 
					continue;
			    project = project.trim();
			    date = date.trim();
			    fileDate.add(date);
			}
			else 
				continue;  // Skip if bad filename
			int sum = 0;
log_debug (String.format("Escribiendo nombre del proyecto (%s) en [0,%d]",project,col));			
			// Escribe el Nombre del Proyecto como encabezado de la sección
			Row r = template_ws.getRow(0);
			Cell c = r.getCell(col);
			c.setCellValue(project);

log_debug("\n@ WorkingOnIt");
			projectSheet = readSheet (filenameIn, "Informe WorkingOnIt");
			if (projectSheet!=null) {
				abstractxSheet = processMatrix (projectSheet);
				template_ws = copyAbstract2Template (abstractxSheet, template_ws, 2, col);
				sum += Integer.parseInt(abstractxSheet[0]);
			}
log_debug("\n@ Stop");			
			projectSheet = readSheet (filenameIn, "Informe Stop");
			if (projectSheet!=null) {
				abstractxSheet = processMatrix (projectSheet);
				template_ws = copyAbstract2Template (abstractxSheet, template_ws, 12, col);
				sum += Integer.parseInt(abstractxSheet[0]);
			}
log_debug("\n@ New");			
			projectSheet = readSheet (filenameIn, "Informe New");
			if (projectSheet!=null) {
				abstractxSheet = processMatrix (projectSheet);
				template_ws = copyAbstract2Template (abstractxSheet, template_ws, 22, col);
				sum += Integer.parseInt(abstractxSheet[0]);
			}
			// Escribe el total de peticiones de las 3 seccioness
			r = template_ws.getRow(1);
			c = r.getCell(col);
			c.setCellValue(sum);

			col += 2;
			numProys++;
		} // End for
		boolean exito = numProys > 0;
		if (!exito) {throw new Exception("Ningun archivo procesado! Todos deben ser '.xls'");}
		date = "";
		String formatedDate = "Fecha: %2s/%2s/%4s";
		// Set date
		if (fileDate.size() != 0) {
			date = fileDate.get(0); // Toma la fecha del 1er archivo
			// Construye la fecha formateada con / / 
			Pattern regExp = Pattern.compile("^(\\d{2})(\\d{2})(\\d{4})$");
			Matcher m = regExp.matcher(date);
			if (m.find())
				formatedDate = String.format(formatedDate,m.group(1),m.group(2),m.group(3)); 
		} else {
			exito = false;
		}
		if (!exito) {throw new Exception("Ningun archivo procesado!");}
		
		// Write formatedDate to cell (0,0)
		template_ws.getRow(0).getCell(0).setCellValue(formatedDate);
		
		// Recalculate formulas (%)
		HSSFFormulaEvaluator.evaluateAllFormulaCells(template_wb);

		// Set Output filename
		filenameOut  =  String.format(filenameOut, date);
		
		// Write to template
		log ("Escribiendo reporte final... " + filenameOut);
		FileOutputStream fos = new FileOutputStream (new File(filenameOut));
		template_wb.write(fos);
		fos.close();
		log ("Proceso terminado con éxito.");
	} // End main

	private static ArrayList<String> getFilenames4Folder (String path, String pattern) throws Exception {
		// pattern ejemplo <string>*<extension> 
		String[] a = pattern.split("\\*"); // Regexp characters must be escaped
		String beg = a[0];
		String end = a[1];
		ArrayList<String> list = new ArrayList<String>();
		final File folder = new File(path);
		for	(final File fileEntry : folder.listFiles()) {
			if (fileEntry.isDirectory()) continue;
			String fn = fileEntry.getName();
			if (!fn.startsWith(beg) || !fn.endsWith(end)) continue;
			list.add(fn);
log_debug ("filename: " + fn);
		}
		return list;
	} // End getFilenames4Folder
	
	/*
	 *  Copy the values of the cells in the source file for
	 *  all the rows and the columns determined by ColsArray;
	 *  copy all them to an array of strings and returns the array
	 */
	private static String[][] readSheet (String fileName, String sheetName) throws Exception {
log_debug ("copiando las celdas...");
		Workbook wb = WorkbookFactory.create(new File(fileName));
		Sheet ws = wb.getSheet(sheetName);
		
		int fileLastRow = ws.getLastRowNum(); // Rows from 0
		// Calcula la última fila del sheet tomando como referencia la Columna 4 del archivo = SUBTAREA
		// Si esta columna es vacía, la fila no debe ser tomada en cuenta
		while (fileLastRow>=FILE_ROW_FIRST && cell2string (ws.getRow(fileLastRow).getCell(FILE_COL_SUBTAREA)).isEmpty()) {
log_debug (fileLastRow + " No es la ultima fila->"+cell2string (ws.getRow(fileLastRow).getCell(ColsArray[FILE_COL_SUBTAREA])));			
			--fileLastRow;
	    }
		if (fileLastRow < FILE_ROW_FIRST) // Si no hay filas que procesar
			return null;
		
		String[][] sheet = new String [fileLastRow-FILE_ROW_FIRST+1][MAT_NUM_COLS];
		for (int i=FILE_ROW_FIRST, row=0; i <= fileLastRow; i++, row++) {
			for (int j=0; j < MAT_NUM_COLS; j++) {
			   sheet[row][j] = cell2string (ws.getRow(i).getCell(ColsArray[j]));
// log_debug ("sheet["+i+","+ColsArray[j]+"]="+sheet[row][j]);
			}
		}
		return sheet;
	} // End readSheet
	
	private static String cell2string(Cell cell) {
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
			result = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			result = cell.getBooleanCellValue();
			break;
		case Cell.CELL_TYPE_ERROR:
			// result = cell.getErrorCellValue();
			result = "";
			break;
		case Cell.CELL_TYPE_BLANK:
			// Blank cell
		case Cell.CELL_TYPE_FORMULA:
			result = "";
			// Will never occur
			break;
		default:	
			throw new RuntimeException ("Qué celda tan extraña !!!");
		}
		return result.toString();
	}
	
	/*
	 * From a <matrix> given values, 
	 * creates a summary and returns it in an array
	 */
	private static String[] processMatrix (String[][] matrix) {
		Integer n = 0; 
		Map<String, Integer> tipoCausa = new HashMap<String, Integer>();
		
		initializeHashes ();
		int numAccionesMitigacion = 0;
		int numObservaciones = 0;
		int rows = matrix.length;
		for (int i=0; i < rows; i++) {
//log_debug ("processMatrix fila ("+i+")");
			// Riesgos
			String risk = matrix[i][COL_RIESGO];
			if (risk != null && risk.length() != 0 && riesgo.get(risk) != null) {
				riesgo.put(risk, riesgo.get(risk)+1);
//log_debug ("   Insertando el riesgo (Col=0)->" + risk);
			}
			
			// Impacto
			String impact = matrix[i][COL_IMPACTO];
			if (impact != null && impact.length() != 0 && impactoEntrega.get(impact) != null)
				impactoEntrega.put(impact, impactoEntrega.get(impact)+1);
			
			// Tipos Causa
			String tcausa = matrix[i][COL_TIPO_CAUSA];
			int val = 0;	
			if (tcausa != null && tcausa.length() != 0) {
				val = tipoCausa.get(tcausa) == null ? 1 : tipoCausa.get(tcausa)+1;
				tipoCausa.put(tcausa, val);
			}
			
			// Acciones de Mitigación
			String mit = matrix[i][COL_ACCIONES_MITIGACION];
			if (mit != null)
				mit = mit.toString();
			if (mit != null && mit.length() != 0)
				++numAccionesMitigacion;
			
				// Observaciones
			String obs = matrix[i][COL_OBSERVACIONES];
			if (obs != null)
				obs = obs.toString();
			if (obs != null && obs.length() != 0)
				++numObservaciones;
			
		} // End for
		// Riesgos + Impactos + Acciones de Mitigación + Observaciones + Tipos Causas
		int rowsResult = riesgo.size() + impactoEntrega.size() + 1 + 1 + 1;
		String[] result = new  String [rowsResult+1]; // +1 por el total de peticiones
		
		int numSubtasks = 0; // Num subtaks with no errors
		int k = 1;
log_debug ("RIESGOS");
		for (Map.Entry<String, Integer> entry : riesgo.entrySet()) {
			n = entry.getValue();
			numSubtasks += n; 
			result[k++] = n.toString();
log_debug (entry.getKey() + ": " + result[k-1]);
		}
log_debug ("IMPACTOS");
		for (Map.Entry<String, Integer> entry : impactoEntrega.entrySet()) {
			result[k++] = entry.getValue().toString();
log_debug (entry.getKey() + ": " + result[k-1]);
		}
		result[k++] = Integer.toString(numAccionesMitigacion);
		result[k++] = Integer.toString(numObservaciones);
log_debug ("CAUSAS");
		String causas = "";
		for (Map.Entry<String, Integer> entry : tipoCausa.entrySet()) {
			causas = causas + entry.getKey() + " (" + entry.getValue().toString() + ")\r\n";
log_debug (entry.getKey() + ": " + entry.getValue().toString());
		}
		result[k] = causas;
		
		// La primera fila del resumen contiene el total de peticiones
		result[0] = Integer.toString(numSubtasks);
log_debug ("Total peticiones: " + result[0]);
        int numErrors = rows - numSubtasks;
log_debug ("Total celdas con error: " + numErrors);

		return result;
	}

	/*
	 * Copies the data in abs 
	 * to template, begining in row, col
	 */
	private static Sheet copyAbstract2Template (String[] abs, Sheet template, int row, int col) {
		int rows = abs.length;
		for (int i=0; i < rows; i++) {
			Row r = template.getRow(row+i);
			Cell c = r.getCell(col);
			// La ultima fila es un String
			if (i == rows-1)
				c.setCellValue (abs[i]);
			else
				c.setCellValue (Integer.parseInt(abs[i]));
		}
		return template;
	}


	/*
	private static void writeSheet (String[][] sheet, String fileName, String sheetName)  throws Exception {
		Workbook wb = new HSSFWorkbook();
	    Sheet sheet1 = wb.createSheet(sheetName);
	    Sheet sheet2 = wb.createSheet("Hoja2");
	    
		int rows = sheet.length;
		int cols = sheet[0].length;
		Cell cell;
		Row rowN;
		
		for (int i=0; i < rows; i++) {
			rowN = sheet1.createRow((short)i);
			for (int j=0; j < cols; j++) {
			    cell = rowN.createCell((short)j);
			    cell.setCellValue(sheet[i][j]); 
			}
		}
		//sheet1.setAutoFilter(new CellRangeAddress (0, rows-1, 0, cols-1));
		
		FileOutputStream fos = new FileOutputStream (new File(fileName));
		// Write the file
		wb.write(fos);
		fos.close();		
	} // End writeSheet
	*/
	
	private static void log (String s) {
		System.out.println(s);
	}
	private static void log_debug (String s) {
		if (DEBUG) 
			System.out.println(s);
	}
	
} // End programa
