package com.indracompany.general;

import java.io.IOException;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

public abstract  class Log {

	static boolean setupDone = false;
	static boolean print = false;
	static boolean toFile = false;
	static Logger logger = null;
	
	public static void setup(boolean prn, String filename) {
		print = prn;
		toFile = !filename.isEmpty();
		if (!toFile) {
		    setupDone = true;
			return; 
		}
		
		logger = Logger.getLogger("MyLog");  
	    FileHandler fh;  
	    System.setProperty("java.util.logging.SimpleFormatter.format", "%1$tY-%1$tm-%1$td %1$tH:%1$tM:%1$tS %4$-6s %2$s %5$s%6$s%n");
	    try {  
	        // This block configure the logger with handler and formatter  
	        fh = new FileHandler(filename);  
	        logger.addHandler(fh);
	        logger.setUseParentHandlers(false);
	        SimpleFormatter formatter = new SimpleFormatter();
	        fh.setFormatter(formatter);
	    } catch (SecurityException e) {  
	        e.printStackTrace();  
	    } catch (IOException e) {  
	        e.printStackTrace();  
	    } 
	    setupDone = true;
	}
	
	public static void print (String msg) throws IOException {
		if (!setupDone) return;
		
		if(print){
			if (toFile){
				logger.info(msg);
			} else { 
				System.out.println ("DEBUG: " + msg);}}
	}
}
