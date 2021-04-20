package com.ltg;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ltg.Application;

public class Application {
	
	static String inFolderPath;
	static String outFolderPath;
	static String outFile;
	static List<HashMap<String, Object>> mapExcelResult;
	
	static int totalModerna;
	static int totalCovovax;
	
	public static void main(String[] args) {
		try {
//			final String infolderPath = args[0];
//			final String outfolderPath = args[1];
			
			String infolderPath = "C:\\Users\\emylyn.audemard\\Documents\\householdConso\\input-err";
			String outfolderPath = "C:\\Users\\emylyn.audemard\\Documents\\householdConso\\err";
			
			String outFileName = "master_HH_Conso_Err.xlsx";
			
			setInFolderPath(infolderPath);
			setOutFolderPath(outfolderPath);
			setOutFile(outFileName);
			
			System.out.println("Running Household Consolidation");
			System.out.print("File Validation....");
			
			if(dirIsEmpty(getInFolderPath())) {//check folder if empty
				System.out.println("Directory is empty.");
				
				System.exit(0);
			}
			
			File directoryPath = new File(getInFolderPath());
			String excelFile[] = directoryPath.list();
			List<List<HashMap<String, Object>>> listAllResult = new ArrayList<List<HashMap<String, Object>>>();
			System.out.println("List of files in the specified directory:");
			
			for(int i=0; i<excelFile.length; i++) {
//				System.out.println(companyNameLookup(excelFile[i]));
				System.out.println(excelFile[i]+" Processing.......");
				
				if(companyNameLookup(excelFile[i]) == "--") {//dont process the file if method return "--", means the file is not white listed
					continue;
				}
				
				listAllResult.add(getSelectedData(excelFile[i])); // all result from excel are stored in Hashmap
				
				setMapExcelResult(listAllResult);
				
				findError(excelFile[i], getMapExcelResult());
			}

			
//			findDuplicateControlNumbers(getMapExcelResult())
			
			
//			for(List<HashMap<String, Object>> r : listAllResult) {
//				System.out.println(r);
//			}
			for(HashMap<String, Object> r : getMapExcelResult()) {
//				System.out.println(r);
			}
			
		}catch (ArrayIndexOutOfBoundsException | IOException e){
	        System.out.println(e);
	    }
	    finally {

	    }


	}
	
	private static void findError(String excelFile, List<HashMap<String, Object>> mapExcelResult) throws FileNotFoundException {
		try {
			Date dNow = new Date();
			SimpleDateFormat ft = new SimpleDateFormat ("yyyy-MM-dd_(hh_mm_ss)");
			
			String[] scn;
    		String companyNamefile = null;
    		if(excelFile.contains("Daily")) {
    			scn = excelFile.split("Daily");
    			
    			companyNamefile = scn[0].trim();
    		}else if(excelFile.contains("Family")) {
    			scn = excelFile.split("Family");
    			
    			companyNamefile = scn[0].replace("_","").trim();
    		}
    		
			
			FileWriter writer = new FileWriter(getOutFolderPath()+"\\"+companyNamefile+"_Reservation_Consolidate_Err_Log_"+ft.format(dNow)+".txt", true);
            BufferedWriter bufferedWriter = new BufferedWriter(writer);
            int totalModerna = 0;
            int totalYestoSwitch = 0;
            
		for(HashMap<String, Object> r : mapExcelResult) {
			
//			System.out.println(r);
//			System.out.println(r.get("switchToCovovax"));
			
			
//			totalModerna += Integer.parseInt((String) r.get("modernaOrders"));
			
//			System.out.println("totalModerna: "+totalModerna);
			
//			if(r.get("switchToCovovax").equals("Yes")) {
//				totalYestoSwitch++;
//			}
			
//			System.out.println("switchToCovovax: "+totalYestoSwitch);
			
			if(r.size() < 6) {
				continue;
			}
			if(!r.isEmpty()) {
				ArrayList<String> errList = new ArrayList<String>();
				
				int cellid = 0;
				
				String companyName;
				
				if(r.get("companyCode").toString() != "--") {
					companyName = r.get("companyCode").toString();
				}else {
					companyName = r.get("companyName").toString();
				}
				
				//============================================================================
				
				if(Integer.parseInt(r.get("covovaxOrders").toString()) != 0 && r.get("CovovaxCtrlNumber").toString() == "--") {
					errList.add("Covovax Reservation Control Number is Blank.");
				}else {
					if(Integer.parseInt(r.get("covovaxOrders").toString()) > 1) {
						if(Integer.parseInt(r.get("covovaxOrders").toString()) != getControlNumberItem(r.get("CovovaxCtrlNumber").toString())) {
							errList.add("Covovax Orders did not match the number of Reservation Control Numbers.");
						}
						
						if(checkCtrlNumberFormat(r.get("CovovaxCtrlNumber").toString())) {
							errList.add("Covovax Reservation Control Number is wrong format. Sample Format: <company code>_<employee number>_C<increment number>.");
						}
						
						if(checkCtrlNumberDelimeter(Integer.parseInt(r.get("covovaxOrders").toString()), r.get("CovovaxCtrlNumber").toString())) {
							errList.add("Covovax  Reservation Control Number is invalid. Control numbers should be separated by comma(,).");
						}
					}
				}
				
				if(Integer.parseInt(r.get("modernaOrders").toString()) != 0 && r.get("ModernaCtrlNumber").toString() == "--") {
					errList.add("Moderna Reservation Control Number is Blank");
				} else {
					if(Integer.parseInt(r.get("modernaOrders").toString()) > 1) {
						if(Integer.parseInt(r.get("modernaOrders").toString()) != getControlNumberItem(r.get("ModernaCtrlNumber").toString())) {
							errList.add("Moderna Orders did not match the number of Reservation Control Numbers.");
						}
						
						if(checkCtrlNumberFormat(r.get("ModernaCtrlNumber").toString())) {
							errList.add("Moderna Reservation Control Number is in wrong format. Sample Format: <company code>_<employee number>_M<increment number>.");
						}
						
						if(checkCtrlNumberDelimeter(Integer.parseInt(r.get("modernaOrders").toString()), r.get("ModernaCtrlNumber").toString())) {
							errList.add("Moderna Reservation Control Number is invalid. Control numbers should be separated by comma(,).");
						}
					}
				}
				
				if(r.get("companyCode").toString() == "--") {
					errList.add("Company Code is Blank.");
				}
				
				if(r.get("employeeNumber").toString() == "--") {
					errList.add("Employee Number is Blank.");
				}
				
				if(r.get("employeeNumber").toString().equals("n/a") || r.get("employeeNumber").toString().equals("na")) {
					errList.add("Invalid Employee Number.");
				}
				
				if(Integer.parseInt(r.get("modernaOrders").toString()) > 40) {
					errList.add("Morderna Orders exceeded order limit.");
				}
				
				if(Integer.parseInt(r.get("covovaxOrders").toString()) > 40) {
					errList.add("Covovax Orders exceeded order limit.");
				}
				
				
				if(errList.size() != 0) {
					System.out.println("ERROR - Row "+r.get("rowNumber")+" "+errList.toString());
					bufferedWriter.write(companyName+" - ERROR - Row "+r.get("rowNumber")+" "+errList.toString());
					bufferedWriter.newLine();
				}
				
			}
			
		}
		
		    bufferedWriter.close();
		    System.out.println("written successfully");
		   
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	      
	      
		
	}

	private static boolean checkCtrlNumberDelimeter(int vOrders, String cn) {
		
		if(cn.contains(",")) {
			String[] controlNumbers = cn.split(",");
			
			if(controlNumbers.length != vOrders) {
				return true;
			}
		}else {
			return true;
		}
		return false;
	}

	private static boolean checkCtrlNumberFormat(String cn) {
		String[] controlNumbers = cn.split(",");
		String[] ctrlFormat;
		
		for ( String s : controlNumbers ) {
	        ctrlFormat = s.split("_");
	        
	        if(ctrlFormat.length != 3) {
	        	return true;
	        }
	    }
	
		return false;
	}

	private static int getControlNumberItem(String cn) {
		String ctrlnumbers[] = cn.split(",");
//		System.out.println(ctrlnumbers.length);
		return ctrlnumbers.length;
		
	}

	private static List<HashMap<String, Object>> getSelectedData(String excelFile) throws FileNotFoundException {
		
		System.out.println("Company Name: "+companyNameLookup(excelFile));
		
		List<HashMap<String, Object>> listsMap = new ArrayList<HashMap<String, Object>>();
		DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
		//check if its director or file
		File file = new File(getInFolderPath()+"\\"+excelFile);
		
		int totalModernaOrders = 0;
		int totalCovovaxOrders = 0;
		
		
		if(!file.isDirectory()) {
		
			List<List<Object>> lists = new ArrayList<List<Object>>();
			
			int record = 0;
			try {
				int totalModerna = 0;
				FileInputStream fis = new FileInputStream(file);
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				XSSFSheet spreadsheet = workbook.getSheetAt(0);
				
				Iterator < Row >  rowIterator = spreadsheet.iterator();
				
				while (rowIterator.hasNext()) {
			    	Row row = rowIterator.next();
			    	
					if (isBlankRow(row)) {
//						System.out.print(isBlankRow(row));
						continue;
					}
			    	
					if(row.getRowNum()==0){
						continue; //just skip the rows if row number is 0, 1, or 2
					}
			    	
			    	Iterator<Cell> cellIterator = row.cellIterator();
			    	List<Object> list = new ArrayList<Object>(); 
			    	HashMap<String, Object> mapList = new HashMap<String, Object>();
			    	
			    	
			    	while (cellIterator.hasNext()) {
			    		Cell cell = cellIterator.next();
			    	
			    		mapList.put("rowNumber", row.getRowNum()+1);
			    		
			    		if(cell.getColumnIndex()==2) { //Completion time
			    			switch (cell.getCellType()) {
				               case NUMERIC:
				            	  mapList.put("completionTime", df.format(cell.getDateCellValue()));
				                  break;
				               case STRING:
					              mapList.put("completionTime", cell.getStringCellValue());
					              break;
				               case BLANK:
				            	   mapList.put("completionTime", "--");
						              break;
							default:
								break;
				            }
			    		}
			    		
			    		
			    		if(cell.getColumnIndex()==18) { //For how many people are you reserving Moderna vaccines?
			    			switch (cell.getCellType()) {
				               case NUMERIC:
				                  mapList.put("modernaOrders", converterStringNum(cell.getNumericCellValue()));
				                  break;
				               case STRING: 
					              mapList.put("modernaOrders", converterStringNum(cell.getStringCellValue()));
					              break;
				               case BLANK:
				            	   mapList.put("modernaOrders", converterStringNum(0));
						              break;
							default:
								break;
				            }
			    		}
			    		
			    		if(cell.getColumnIndex()==19) { //If your Moderna quantity is not available, would you like to switch this quantity from Moderna to Covovax (Novavax)?
			    			switch (cell.getCellType()) {
				               case STRING:
						        	 String res;
						        	 if(cell.getStringCellValue().toString().contains("No")) {
						        		 res = "No";
						        	 } else {
						        		 res = "Yes";
						        	 }
				            	  mapList.put("switchToCovovax", res);
				                  break;
				               case NUMERIC:
				            	   mapList.put("switchToCovovax", cell.getNumericCellValue());
					              break;
				               case BLANK:
				            	   mapList.put("switchToCovovax", "--");
						              break;
							default:
								break;
				            }
			    		}
			    	
			    		
			    		if(cell.getColumnIndex()==21) { //For how many people are you reserving Moderna vaccines?
			    			switch (cell.getCellType()) {
				               case NUMERIC:
				                  mapList.put("covovaxOrders", converterStringNum(cell.getNumericCellValue()));
				                  break;
				               case STRING: 
					              mapList.put("covovaxOrders", converterStringNum(cell.getStringCellValue()));
					              break;
				               case BLANK:
				            	   mapList.put("covovaxOrders", converterStringNum(0));
						              break;
							default:
								break;
				            }
			    		}
			    		
			    		if(cell.getColumnIndex()==24) { //Company Code
			    			switch (cell.getCellType()) {
				               case STRING:
				            	   mapList.put("companyCode", cell.getStringCellValue());
				                  break;
				               case NUMERIC:
				            	   mapList.put("companyCode", cell.getNumericCellValue());
					              break;
				               case BLANK:
				            	   mapList.put("companyCode", "--");
				            	   System.out.println("blank");
						              break;
							default:
								mapList.put("companyCode", "--");
								System.out.println("def");
								break;
				            }
			    			
			    		
			    		}
			    		
			    		if(cell.getColumnIndex()==25) { //Moderna Control Number
			    			switch (cell.getCellType()) {
				               case STRING:
				            	   mapList.put("ModernaCtrlNumber", cell.getStringCellValue());
				                  break;
				               case NUMERIC:
				            	   mapList.put("ModernaCtrlNumber", cell.getNumericCellValue());
					              break;
				               case BLANK:
				            	   mapList.put("ModernaCtrlNumber", "--");
						              break;
							default:
								break;
				            }
			    		}
			    		
			    		if(cell.getColumnIndex()==26) { //Covovax Control Number
			    			switch (cell.getCellType()) {
				               case STRING:
				            	   mapList.put("CovovaxCtrlNumber", cell.getStringCellValue());
				                  break;
				               case NUMERIC:
				            	   mapList.put("CovovaxCtrlNumber", cell.getNumericCellValue());
					              break;
				               case BLANK:
				            	   mapList.put("CovovaxCtrlNumber", "--");
						              break;
							default:
								break;
				            }
			    		}
			    		
			    		if(cell.getColumnIndex()==10) { //employeeNumber
			    			switch (cell.getCellType()) {
				               case STRING:
				            	   mapList.put("employeeNumber", cell.getStringCellValue().toLowerCase().replaceAll(" ", ""));
				                  break;
				               case NUMERIC:
				            	   mapList.put("employeeNumber", cell.getNumericCellValue());
					              break;
				               case BLANK:
				            	   mapList.put("employeeNumber", "--");
						              break;
							default:
								break;
				            }
			    		}
			    		
			    		String[] scn;
			    		String companyName = null;
			    		if(excelFile.contains("Daily")) {
			    			scn = excelFile.split("Daily");
			    			
			    			companyName = scn[0].trim();
			    		}else if(excelFile.contains("Family")) {
			    			scn = excelFile.split("Family");
			    			
			    			companyName = scn[0].replace("_","").trim();
			    		}
			    		
			    		mapList.put("companyName", companyName);

			    	}
			    	
//			    	if(Integer.parseInt(mapList.get("modernaOrders").toString()) != 0 && Integer.parseInt(mapList.get("covovaxOrders").toString()) != 0) {
//			    		listsMap.add(mapList);
//			    	}
			    	listsMap.add(mapList);
			    	
				}
				
				for(HashMap<String, Object> s : listsMap) {
//					System.out.println(s);
				}
				
				System.out.println();
				System.out.println();
				
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}else {
			System.out.println("--- IS A DIRECORY");
		}
	
		return listsMap;
	}

	private static Object companyNameLookup(String excelFile) {
		String[] scn;
		String companyName = null;
		
		if(excelFile.contains("Daily")) {
			scn = excelFile.split("Daily");
			
			companyName = scn[0].trim();
		}else if(excelFile.contains("Family")) {
			scn = excelFile.split("Family");
			
			companyName = scn[0].replace("_","").trim();
		}
		
		
		HashMap<String, String> cc = new HashMap<String, String>();
		
		cc.put("ALL", "All Seasons Realty Corp.");
		cc.put("APL", "Allianz-PNB Life Insurance, Inc. (APLII)");
//		cc.put("APL", "Allianz-PNB Life Insurance, Inc.");
		cc.put("ABI", "Asia Brewery, Inc. (ABI) and Subsidiaries");
//		cc.put("ABI", "ABI, its Subsidiaries, and Affiliates"); // 
		cc.put("BCH", "Basic Holdings Corp.");
		cc.put("CPH", "Century Park Hotel");
		cc.put("EPP", "Eton Properties Philippines, Inc. (Eton) and Subsidiaries");
//		cc.put("EPP", "EPPI and its Subsidiaries");
		cc.put("FFI", "Foremost Farms, Inc.");
		cc.put("FTC", "Fortune Tobacco Corp.");
		cc.put("GDC", "Grandspan Development Corp.");
		cc.put("HII", "Himmel Industries, Inc.");
		cc.put("LRC", "Landcom Realty Corp.");
		cc.put("LTG", "LT Group, Inc. (Parent Company)");
		cc.put("LTGC", "LTGC Directors");
//		cc.put("LTGC", "LT Group of Companies Directors");
//		cc.put("MAC", "MacroAsia Corp., Subsidiaries & Affiliates");
//		cc.put("MAC", "MAC, its Subsidiaries, and Affiliates");
		cc.put("MAC", "MacroAsia Corp., Subsidiaries & Affiliates");
		cc.put("PAL", "Philippine Airlines, Inc. (PAL), Subsidiaries and Affiliates");
//		cc.put("PAL", "PAL, its Subsidiaries, and Affiliates");
		cc.put("PNB", "Philippine National Bank (PNB) and Subsidiaries");
//		cc.put("PNB", "PNB and its Subsidiaries");
		cc.put("PMI", "PMFTC Inc.");
		cc.put("RAP", "Rapid Movers & Forwarders, Inc.");
		cc.put("TYK", "Tan Yan Kee Foundation, Inc. (TYKFI)");
		cc.put("TDI", "Tanduay Distillers, Inc. (TDI) and subsidiaries");
//		cc.put("TDI", "TDI, its Subsidiaries, and Affiliates");
		cc.put("CHI", "Charter House Inc.");
//		cc.put("SPV", "Grandholdings Investments (SPV-AMC), Inc.");
//		cc.put("SPV", "Opal Portfolio Investments (SPV-AMC), Inc.");
		cc.put("SPV", "SPV-AMC Group");
		cc.put("SPV", "SPV Group");
		cc.put("TMC", "Topkick Movers Corporation");
		cc.put("UNI", "University of the East (UE)");
		cc.put("UER", "University of the East Ramon Magsaysay Memorial Medical Center (UERMMMC)");
//		cc.put("UER", "UERMMMC");
		cc.put("VMC", "Victorias Milling Company, Inc. (VMC)");
		cc.put("ZHI", "Zebra Holdings, Inc.");
		cc.put("STN", "Sabre Travel Network Phils., Inc.");
		cc.put("TMC", "Topkick Corp.");
//		cc.put("TMC", "Topkick Movers Corporation");
		
//		if(!cc.containsKey(excelFile)) {
//			return "--";
//		}
		
//		return companyName;
		
//		if(getKey(cc, companyName) == null) {
//			return "--";
//		}
//		return getKey(cc, companyName) +" - "+ companyName;
		
		
		return getKey(cc, companyName);
	}
	
	public static <K, V> K getKey(Map<K, V> map, V value) {
        for (Map.Entry<K, V> entry : map.entrySet()) {
            if (value.equals(entry.getValue())) {
                return entry.getKey();
            }
        }
        return null;
    }

	private static Object converterStringNum(Object numericCellValue) {
		double d = Double.valueOf(numericCellValue.toString()).doubleValue();
		int orders = (int)d;
		
		
		return String.valueOf(orders);
	}

	private static boolean checkFile(String fileformat, String file) {
		File f = new File(file);
		if(f.isFile() && !f.isDirectory()) { 
			
			String filename = f.getName().toLowerCase();
			
			if(!filename.endsWith(fileformat)) {
				System.out.println(file + " is not valid excel format.");
				
				return false;
			}
		}else {
			System.out.println(file + " does not exist.");
			
			return false;
		}
		
		return true;
	}

	private static boolean dirIsEmpty(String inFolderPath) throws IOException {
		Path p = Paths.get(inFolderPath);
		
	    if (Files.isDirectory(p)) {
	        try (Stream<Path> entries = Files.list(p)) {
	            return !entries.findFirst().isPresent();
	        }
	    }
	        
	    return false;
	}

	public static String getInFolderPath() {
		return inFolderPath;
	}

	public static void setInFolderPath(String inFolderPath) {
		Application.inFolderPath = inFolderPath;
	}

	public static String getOutFile() {
		return outFile;
	}

	public static void setOutFile(String outFile) {
		Application.outFile = outFile;
	}

	public static List<HashMap<String, Object>> getMapExcelResult() {
		return mapExcelResult;
	}

	public static void setMapExcelResult(List<List<HashMap<String, Object>>> allResult) {
		List<HashMap<String, Object>> lists = new ArrayList<HashMap<String, Object>>();
		
		for(List<HashMap<String, Object>> s : allResult) {
			for(HashMap<String, Object> r : s) {
				
//				System.out.println(r);
				lists.add(r);
				

    		}
		}
		
		Application.mapExcelResult = lists;
	}

	public static String getOutFolderPath() {
		return outFolderPath;
	}

	public static void setOutFolderPath(String outFolderPath) {
		Application.outFolderPath = outFolderPath;
	}

	public static int getTotalModerna() {
		return totalModerna;
	}

	public static void setTotalModerna(int totalModerna) {
		
		Application.totalModerna += totalModerna;
	}

	public static int getTotalCovovax() {
		return totalCovovax;
	}

	public static void setTotalCovovax(int totalCovovax) {
		Application.totalCovovax += totalCovovax;
	}
	
	private static boolean isBlankRow(Row row) {
        Cell cell;
        boolean result = true;
       
        for(int col = 0; col <= 72; col ++) {               
            cell = row.getCell(col);       
            /*if(row.getRowNum()>=8400) {
                System.out.println(cell + " - " + isCellEmpty(cell, false) );
            }*/
            if(!isCellEmpty(cell, false)) {
                result = false;
                break;                       
            }
        }
        return result;
    }
	
	 private static boolean isCellEmpty(Cell cell, boolean checkForZero) {       
	        if (cell == null) {
	            return true;
	        }   
	        if (cell.getCellType() == CellType.BLANK) {
	            return true;
	        }   
	        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) {
	            return true;
	        }   
	        if (checkForZero && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == 0) {
	            return true;
	        }
	        if (cell.getCellType() == CellType.FORMULA) {
	            CellType cellType = cell.getCachedFormulaResultType();
	            if(cellType == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) {                                                       
	                return true;                                                   
	            }
	           
	            if(checkForZero && cellType == CellType.NUMERIC && cell.getNumericCellValue() == 0) {                                                       
	                return true;                                                   
	            }
	        }
	        return false;
	   }

}
