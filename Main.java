import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.concurrent.TimeUnit;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main 
{
	private static XSSFWorkbook wb;

	public static void main(String args[]) throws IOException, InterruptedException 
	{
		JOptionPane.showMessageDialog(null ,"Please select the tax file.");
		String taxfilename = PickAFile();
		System.out.println(taxfilename);
		JOptionPane.showMessageDialog(null ,"Please select the folder with the pay files.");
		String payfilename = PickAPayFile();
		String mergedfilename = MergeTextFiles(payfilename);
		System.out.println(payfilename);
		OpenInExcel(mergedfilename, taxfilename, payfilename);
		
	} 

	public static String PickAPayFile()
	{
	    JFileChooser chooser = new JFileChooser();
	    
	    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
	    
	    chooser.setDialogTitle("Please select the folder that contains all the pay files.");;
	    
	    int returnVal = chooser.showOpenDialog(null);
	    
	    if(returnVal == JFileChooser.APPROVE_OPTION) 
	    	{
	    		System.out.println("You chose this file: " + chooser.getSelectedFile().getName());
	        	return chooser.getSelectedFile().getAbsolutePath();
	    	}
	    else
	    	System.exit(0);
	    	return null;
	}
	
	public static String PickAFile()
	{
	    JFileChooser chooser = new JFileChooser();
	    
	    chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
	    
	    chooser.setDialogTitle("Please select the file that contains all the tax files.");;
	    
	    int returnVal = chooser.showOpenDialog(null);
	    
	    if(returnVal == JFileChooser.APPROVE_OPTION) 
	    	{
	    		System.out.println("You chose this file: " + chooser.getSelectedFile().getName());
	        	return chooser.getSelectedFile().getAbsolutePath();
	    	}
	    else
	    	System.exit(0);
	    	return null;

	}
	
	public static String MergeTextFiles(String filelocation) throws IOException, InterruptedException
	{
		ProcessBuilder builder = new ProcessBuilder("cmd.exe", "/c", "cd " + filelocation + " && copy *txt mergedfiles.txt");
	        
		builder.redirectErrorStream(true);
	        
	    Process p = builder.start();
	    
	    TimeUnit.SECONDS.sleep(3);
	        
	    return filelocation + "\\mergedfiles.txt";
	}
	
	public static void OpenInExcel(String mergedpayfilelocation, String taxfilelocation, String filelocation) throws IOException
	{
		
		  LinkedList<String[]> text_lines = new LinkedList<>();
		    try (BufferedReader br = new BufferedReader(new FileReader(mergedpayfilelocation))) 
		    {
		        String sCurrentLine;
		        while ((sCurrentLine = br.readLine()) != null) 
		        {
		        	String [] tCurrentLine = sCurrentLine.split("\\s" + "\\s" + "\\s");
		        	for(int x = 0; x < tCurrentLine.length; x++)
		        	{
		        		
		        		if (tCurrentLine[x].length() == 4)
		        		{
		        			if(tCurrentLine[x].contains("ARMS"))
		        			{
		        				
		        			}
		        			else
		        			{
		        			tCurrentLine[x-1] = "!";
		        			}
		        		}
		        		else if (tCurrentLine[x].contains("Total Tickets"))
		        		{
		        			tCurrentLine[x] = "!";

		        			tCurrentLine[x+3] = "!";
		        		}
		        		else if (tCurrentLine[x].contains("Total Payment"))
		        		{
		        			tCurrentLine[x] = "!";

		        			tCurrentLine[x+3] = "!";
		        		}
		        	
		        	}
		        	
		        	text_lines.add(tCurrentLine);                 
		        }
		    } catch (IOException e) 
		    {
		        e.printStackTrace();
		    }
		    
		    LinkedList<String[]> text_linez = new LinkedList<>();
		    try (BufferedReader br = new BufferedReader(new FileReader(taxfilelocation))) 
		    {
		        String sCurrentLine;
		        while ((sCurrentLine = br.readLine()) != null) 
		        {
		        	String [] tCurrentLine = sCurrentLine.split("$");
		        	text_linez.add(tCurrentLine);                 
		        }
		    } catch (IOException e) 
		    {
		        e.printStackTrace();
		    }
		    
		// n = JOptionPane.showInputDialog("What would you like to save the file as?", null);
	    
	    XSSFWorkbook workbook = new XSSFWorkbook();
	    XSSFSheet sheet = workbook.createSheet("Pay Files");
	    XSSFCellStyle stylepay = workbook.createCellStyle();
	    XSSFCellStyle styletax = workbook.createCellStyle();
	    stylepay.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
	    styletax.setFillForegroundColor(IndexedColors.RED1.getIndex());
	    stylepay.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    styletax.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	    int row_num = 0;
	    
	    for(String[] line : text_lines)
	    {
	        Row row = sheet.createRow(row_num++);
	        int cell_num = 0;
	        
	        for(String value : line)
	        {
	            Cell cell = row.createCell(cell_num++);
	            cell.setCellValue(value);
	            cell.setCellStyle(stylepay);
	        }
	        
	    }
	    
	    XSSFSheet s = workbook.createSheet("Tax Files");
	    XSSFSheet c = workbook.createSheet("Combo");
	    XSSFSheet f = workbook.createSheet("Tax Files with no Pay");

	   
	    row_num = 0;
	    
	    for(String[] line : text_linez)
	    {
	        Row row = s.createRow(row_num++);
	        int cell_num = 0;
	        
	        for(String value : line)
	        {
	            Cell cell = row.createCell(cell_num++);
	            cell.setCellValue(value);
	            cell.setCellStyle(styletax);
	        }
		
	    }
	    
	    ArrayList<String> Tax = new ArrayList<String>();
	    
	    for (Row row : s)
	    {
	    	for (Cell cell : row)
	    	{
	    		//cell.getStringCellValue().substring(cell.getStringCellValue().indexOf("$") + 19, cell.getStringCellValue().indexOf("$") + 21).equals("$5$")
	    		if(cell.getStringCellValue().contains("$5$"))
	    		{
	    			Tax.add(cell.getStringCellValue().substring(0, cell.getStringCellValue().indexOf("$")));
	    			Tax.add(cell.getStringCellValue().substring(cell.getStringCellValue().indexOf("$") + 12, cell.getStringCellValue().indexOf("$") + 17) + "." + cell.getStringCellValue().substring(cell.getStringCellValue().indexOf("$") + 17, cell.getStringCellValue().indexOf("$") + 19));
	    		}
	    	}
	    }
	    
	    row_num = 0;
	    int y = 0;
	    while (y < Tax.size())
	    {
	    	Row row = c.createRow(row_num++);
	    	
	    	for (int x = 0; x <2; x++)
	    	{
	    		if(y+x < Tax.size()) {
	    			if(Tax.get(x).equals(" "))
	    			{
	    				
	    			}
	    			else
	    			{
	    		Cell cell = row.createCell(x);
	            cell.setCellValue(Double.parseDouble(Tax.get(y+x)));
	            cell.setCellStyle(styletax);
	            }}
	    	}
	    	
	    y=y+2;
	     
	    }
	    
	    ArrayList<String> Pay = new ArrayList<String>();
	    
	    for (Row row : sheet) 
	    {
	      for (Cell cell : row) 
	      {
	    	  if (cell.getStringCellValue().equals("!"))
	    	  {
	    		  
	    	  }
	    	  else if(cell.getStringCellValue().contains("Law Firm Billing"))
	    	  {
	    		  
	    	  }
	    	  else if(cell.getStringCellValue().contains(""))
	    	  {
	    		  
	    	  }
	    	  else
	    	  {
	    		  Pay.add(cell.getStringCellValue());
	    	  }
	      }
	    }
	    
	    for(int x = 0; x < Pay.size(); x++)
	    {
	    	if(Pay.get(x).equals(" "))
	    	{
	    		Pay.remove(x);
	    		
	    	}
	    	if(Pay.get(x).equals(""))
	    	{
	    		Pay.remove(x);
	    	}
	    	if(Pay.get(x).contains("Tckt Num"))
	    	{
	    		Pay.remove(x);
	    		
	    	}
	    	if(Pay.get(x).contains("Pymt Amt"))
	    	{
	    		Pay.remove(x);
	    		
	    	}
	    	if(Pay.get(x).contains("ARMS"))
	    	{
	    		Pay.remove(x);
	    		
	    	}
	    	if(Pay.get(x).contains("REG"))
	    	{
	    		Pay.remove(x);
	    		
	    	}
	    	if(Pay.get(x).contains("PYMTPLN"))
	    	{
	    		Pay.remove(x);
	    		
	    	}
	    }
	    

	    
	    y = 0;
	    Pay.add(0, " ");
	    while (y < Pay.size())
	    {
	    	Row row = c.createRow(row_num++);
	    	
	    	for (int x = 0; x <2; x++)
	    	{
	    		if(y+x <Pay.size()) {
	    			try {
	    		Cell cell = row.createCell(x);
	            cell.setCellValue(Double.parseDouble(Pay.get(y+x)));
	            cell.setCellStyle(stylepay);}
	    			catch(NumberFormatException ex) {}}
	    	}
	    	
	    y=y+2;
	     
	    }
	    
	    {
			
			XSSFSheet l = workbook.getSheetAt(2);
			XSSFSheet m = workbook.getSheetAt(3);
			
		    String n = JOptionPane.showInputDialog("What would you like to save the file as?", null);
		    String cFileName = filelocation + "/" + n + ".xlsx";
		
		    ArrayList<Double> unmatched = new ArrayList<Double>();
		    int cellLocation = 0;
		    double currentValue = 0;
		    
		    for (Row row : l)
		    {
		    	for (Cell cell : row)
		    	{
		    		if (cellLocation <= 1)
		    		{
		    			if(row.getRowNum() > 0)
		    			{
		    				currentValue = cell.getNumericCellValue();
		    			}
		    			if(cell.getCellStyle().getFillForegroundColor() == IndexedColors.RED1.getIndex())
		    			{
		    				unmatched.add(currentValue);
		    			}
		    			cellLocation++;
		    		}
		    		
		    	}
		    	cellLocation = 0;
		    }
		    
		    for (Row row : l)
		    {
		    	for (Cell cell : row)
		    	{
		    		if (cellLocation == 0)
		    		{
		    			if(row.getRowNum() > 0)
		    			{
		    				currentValue = cell.getNumericCellValue();
		    			}
		    			if(cell.getCellStyle().getFillForegroundColor() == IndexedColors.BRIGHT_GREEN.getIndex())
		    			{
		    				if(unmatched.contains(currentValue))
		    				{
		    					int temp = unmatched.indexOf(currentValue);
		    					unmatched.remove(temp+1);
		    					unmatched.remove(temp);
		    				}
		    			}
		    			cellLocation++;
		    		}
		    		
		    	}
		    	cellLocation = 0;
		    }
		    
		    int counter = 0;
		    int row_n = 0;
		    
		    while (counter < unmatched.size())
		    {
		    	Row row = m.createRow(row_n++);
		    	
		    	for (int x = 0; x <2; x++)
		    	{
		    		Cell cell = row.createCell(x);
		            cell.setCellValue(unmatched.get(counter+x));
		            cell.setCellStyle(styletax);
		    	}
		    	
		    counter=counter+2;
		     
		    }
		    
		    
		    
		    try 
		    {
		        FileOutputStream o = new FileOutputStream(cFileName);
		        workbook.write(o);
		        o.close();
		    } 
		    catch (FileNotFoundException ex) 
		    {
		    	System.out.println("Failed.");
		    	System.exit(0);
		    }
		}
	    
	
	}
}
