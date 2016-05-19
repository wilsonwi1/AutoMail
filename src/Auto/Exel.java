package Auto;
/*
 * realworldretail
 * William Wilson
 * 05/11/2015
 */
import java.io.*;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.*;
/*
 *   Extracts excel sheet info
*/
public class Exel {
	private static String exelFile;
	private static XSSFWorkbook excelBook;
	private static int noRows,noCC;  
	//Config parameters 
	private XSSFSheet configSheet;
	private static String tabCmdLoc,server,tabUser,pass,siteName,file,param1Name,param2Name;
	private static String pdf,repLoc,repName,fabLogin,from,sub,logSuccess,logFail;
	//Email parameters
	private static XSSFSheet emailSheet;
	private static String param1,param2,identifier,txt,emailTO,emailCC[];
	private String fabooti = "\"C:\\Program Files (x86)\\Febooti Command line email\\febootimail.exe\" ";
	//Formatter
	private static DataFormatter formatter = new DataFormatter();
	
	Exel(String fileName, int numberCCs){
		exelFile = Location()+ fileName;
		try{		
			excelBook = new XSSFWorkbook(new FileInputStream(exelFile));
			configSheet = excelBook.getSheet("Config");
		}
		catch(IOException e){
			System.out.println("Error reading exel file");
		}
		//System.out.println(exelFile);
		noRows = 2;
		noCC = numberCCs;		
	}
	/*
	 *   Reads configuration details
	*/
	public void readCofig(){
		try {
			excelBook = new XSSFWorkbook(new FileInputStream(exelFile));
			configSheet = excelBook.getSheet("Config");
			int r = 0;
			tabCmdLoc =readCell(configSheet,r++,1);
		    server = readCell(configSheet,r++,1);
		    tabUser = readCell(configSheet,r++,1);
		    pass = readCell(configSheet,r++,1);
		    siteName = readCell(configSheet,r++,1);
		    file = readCell(configSheet,r++,1);
		    param1Name = readCell(configSheet,r++,1);
		    param2Name = readCell(configSheet,r++,1);
		    pdf = readCell(configSheet,r++,1);
		    repLoc = readCell(configSheet,r++,1);
		    repName = readCell(configSheet,r++,1);
		    fabLogin = readCell(configSheet,r++,1);
		    from = readCell(configSheet,r++,1);	
		    sub = readCell(configSheet,r++,1);	
		    logSuccess = readCell(configSheet,r++,1);
		    logFail = readCell(configSheet,r++,1);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/*
	 *   Reads given cell
	*/
	private static String readCell(XSSFSheet Sheet, int r, int c){
		String var = formatter.formatCellValue( Sheet.getRow(r).getCell(c));
    	if(var==""){	
    		System.out.println(Sheet.getSheetName()+"("+r+","+c+"):Cell Empty");
    	}
    	return var;
	}
	/*
	 *   Gets exact number of valid rows / reports to be emailed
	*/
	public int getNoRows() throws FileNotFoundException, IOException{
		emailSheet = excelBook.getSheet("Email");
		noRows =  emailSheet.getPhysicalNumberOfRows();
	    int i ;
	    for(i = 1; i < noRows; i++){
	    	XSSFCell cell = (XSSFCell) emailSheet.getRow(i).getCell(0);
	    	if(cell == null){
				break;
			}
	    	String temp = formatter.formatCellValue(cell); 
	    	if(temp.trim().length() == 0){
				break;
			}				
	    }	
	    noRows= i;
	    return noRows;
	}
	/*
	 *   Reads specified row from excel sheet
	*/
	public static boolean ReadNextRow(int row) throws FileNotFoundException, IOException{
		XSSFRow r = emailSheet.getRow(row);
	    if(r !=null){
	    
	    	XSSFCellStyle xx = r.getCell(0).getCellStyle();
	    	if(xx.getFillForegroundColor()==XSSFCellStyle.NO_FILL) {
	    		System.out.println("Fill");
	    		return false;
	    	}
	    	String val = formatter.formatCellValue(r.getCell(0));
	    	if(val!=""){
	    		param1 = val;
	    		if(param2Name!=""){
	    			param2 = readCell(emailSheet,row,1);
	    		}
	    		identifier = readCell(emailSheet,row,2);
			    txt = readCell(emailSheet,row,3);
				emailTO = readCell(emailSheet,row,4);
				emailCC = new String[noCC];								//New blank array of cc emails	
				for(int j=0;j<noCC;j++){
					XSSFCell cell = (XSSFCell) r.getCell(j+5);
					if((cell !=null)){
						String temp = r.getCell(j+5).getStringCellValue();
						temp = temp.trim();
						if(temp.length() != 0){
							emailCC[j] = temp;
						}
					}
				}					
	    	}
		}
	    return true;
	}
	/*
	 *   Finds this java file location
	*/
	public static String Location(){
		String loc = AutoEmail.class.getProtectionDomain().getCodeSource().getLocation().toString();
        String fName = new java.io.File(AutoEmail.class.getProtectionDomain().getCodeSource().getLocation().getPath()).getName();
        if(loc.contains(".jar")){
        	loc = loc.replace(fName, "");					//delete AutoMail.jar off the end
        }
        loc = loc.replaceFirst("file:/", "");
        loc = loc.replaceFirst("%20", " ");
        loc = loc.replace('/', '\\');
		return loc;
	}
	/*
	 *   Getters
	*/
	public String getReport(){
		String name =(repName + identifier).replace(" ", "_");
		return repLoc +  name + ".pdf";
	}
	public String getTabLoginCmd(){
		String cmd = tabCmdLoc + "tabcmd login -s " + server + " -u " + tabUser
				+ " --password-file " + pass + " --no-certcheck -t \"" + siteName + "\"";	
		return cmd;
	}
	private String getRepReq(){	
		if(param2Name==""){
			return "\"" + file + "?" + param1Name + "=" + param1 + "\"";
		}
		return "\"" + file + "?" + param1Name + "=" + param1 + "&" + param2Name + "=" + param2+ "\"";
	}
	public String getTabRepReqCmd(){
		return tabCmdLoc + "tabcmd export " + getRepReq()
				+ " --" + pdf +" --Timeout 3600 --no-certcheck -f \"" 
				+ getReport()+"\"";
	}
	public String getFabCmd(){
		String cmd = fabooti + fabLogin + " -FROM "+ from + " -TO " + emailTO;
		for(int j=0;j<noCC;j++){
			if(emailCC[j]==null){
				break;
			}
			cmd += " -CC "+emailCC[j];	
		}
		cmd += 	" -TEXT \"" + txt + "\" " 
				+" -ATTACH \"" + getReport() + "\" "
				+" -SUBJECT \"" + sub +" "+identifier+"\" "
				+"-LS \""+logSuccess+"\"  " + "-LF \""+logFail+"\" ";	
		return cmd;
	}
	public String getParam1(){
		return param1;
	}
	public String getParam2(){
		return param2;
	}
}