import java.io.File;
import java.util.*;
import java.util.Scanner;
import java.util.Map.Entry;
import java.io.IOException;
import java.io.FileWriter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.FileOutputStream;
import java.io.FileInputStream;
class Utility 
{
	public static HashMap< ArrayList<String>,ArrayList<String> >map_notfound=null;
	public static HashMap< ArrayList<String>,ArrayList<String> >map_modified=null;
	public static HashMap< ArrayList<String>,ArrayList<String> >map_found=null;
	public static void writeFile(String FileName,String arg,boolean parameter)
	{
		try
		{
	        FileWriter writer = new FileWriter(FileName, parameter);
	        writer.write(arg.toString());
	        writer.write("\r\n");   // write new line
	        writer.close();
	    } 
		catch (IOException e) 
		{
	        e.printStackTrace();
		}
	}
	private static String getFileExtension(String fileName) 
	{
        fileName = fileName.toString();
        if(fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
        	return fileName.substring(fileName.lastIndexOf(".")+1);
        else return "";
    }
	public static ArrayList<String>getMainFolders(String path) 
	{
		ArrayList<String>mainFolders=new ArrayList<String>();
	
        File root = new File( path );
        File[] list = root.listFiles();
        if(list==null)
    		return null;
        for(File f:list)
        {
        	mainFolders.add(f.toString());
        }        
      return mainFolders;
    }
	public static void getPath(String path,ArrayList<String>pathsXlsx,ArrayList<String>pathsXls) 
	{
		//System.out.println(path.toString());
        File root = new File( path );
        File[] list = root.listFiles();
      //  System.out.println(path+list.length);
        if(list==null)
        		return;
        for ( File f : list ) 
        {
            if ( f.isDirectory() ) 
            {
                getPath( f.getAbsolutePath().toString(),pathsXlsx,pathsXls);
            }
            else 
            {
            	String s=f.getAbsolutePath();
            	String fileExtn = Utility.getFileExtension(s);
                if(fileExtn.equals("xlsx") ) 
                {	
                	pathsXlsx.add(s);
                }
                else if(fileExtn.equals("xls"))
                {
                	pathsXls.add(s);
                }
            }           
        }
    }
	public static String stripExtension(final String path)
	{
	    return path != null && path.lastIndexOf(".") > 0 ? path.substring(0, path.lastIndexOf(".")) : path;
	}
	public static boolean match(ArrayList<String> tempExcelLinked, ArrayList<String>Utility )
	{
		if(!(stripExtension(tempExcelLinked.get(tempExcelLinked.size()-1)).equals(stripExtension(Utility.get(Utility.size()-1)))))
			return false;
		for(int i=0;i<tempExcelLinked.size()-1;i++)
		{
			if(!(tempExcelLinked.get(i).toLowerCase().equals(Utility.get(i).toLowerCase())))
			{
				return false;
			}
		}
		return true;
	}
	public static void readExcel_xlsx(String mainFolder,String absolutePathExtracted, ArrayList<String>relPathExtracted)
    {
        try
        {
        	File excel=new File(absolutePathExtracted);
            FileInputStream file = new FileInputStream(excel);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            if(sheet==null)
            {
            	workbook.close();
            	return;
            }
            Row row = sheet.getRow(19);
            if(row==null)
            {
            	workbook.close();
            	return;
            }
            Iterator<Cell> cellIterator = row.cellIterator();
            int colNo=0;
            boolean flag=false;
            String excelLinked=new String();
            while (cellIterator!=null && cellIterator.hasNext())
            {
                Cell cell = cellIterator.next();
                switch (cell.getCellType())
                {
                    case Cell.CELL_TYPE_STRING:
                    	excelLinked=cell.getStringCellValue();
                        if(excelLinked.contentEquals("EXCEL TO BE LINKED"))
                        {
                        	colNo=(cell.getColumnIndex());
                        	flag=true;
                        	break;
                        }
                    break;
                    case Cell.CELL_TYPE_NUMERIC:
                    	break;
                    default:
                    	break;       
                }
                if(flag)
                	break;
            }
            for (int rowIndex = 20; flag==true && rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            	  row = sheet.getRow(rowIndex);
            	  if (row != null) {
            	    Cell cell = row.getCell(colNo);
            	    if (cell != null && cell.getCellType()==Cell.CELL_TYPE_STRING && row.getCell(colNo).getStringCellValue().length()!=0) 
            	    {
            	    	   boolean flag1=false;
	            	       excelLinked=(row.getCell(colNo).getStringCellValue()).toString();
	            	       StringTokenizer st = new StringTokenizer(excelLinked.toString(),"\\//");  
        	    		   ArrayList<String>listExcelLinked=new ArrayList<String>();
        	    		   while (st.hasMoreTokens()) 
        			       {  
        			         listExcelLinked.add(st.nextToken());
        			       }
        	    		   int size=listExcelLinked.size();
        	    		   if(size==0)
        	    			   continue;
        	    		   StringTokenizer st1=null;
        	    		   ArrayList<String>list1=new ArrayList<String>();
        	    		   if((Utility.map_found).containsKey(listExcelLinked))
        	    		   {
        	    			  writeFile("CorrectPaths.txt",absolutePathExtracted+"     -     "+excelLinked.toString()+"\n",true);
          	    			  System.out.println("CorrectPath"+"    -    "+excelLinked);
          	    			  continue;
        	    		   }
        	    		   else if((Utility.map_modified).containsKey(listExcelLinked))
        	    		   {
        	    			   list1=map_modified.get(listExcelLinked);
        	    			   String updatedExcelLinked=excelLinked;
        	    			   for(int i=0;i<size;i++)
        	    			   {
        	    				   if(!(listExcelLinked.get(i).equals(list1.get(i))))
        	    				   {
        	    					   updatedExcelLinked=excelLinked.replace(listExcelLinked.get(i), list1.get(i));
        	    				   }
        	    			   }
        	    			   cell.setCellValue(updatedExcelLinked);
           	    			    FileOutputStream os=new FileOutputStream(excel);
           	    			    workbook.write(os);
           	    			    os.close();
           	    			  	writeFile("LinksModified.txt",absolutePathExtracted+"	 -	"+excelLinked+"	 -	"+updatedExcelLinked.toString()+"\n",true);
	           	    			System.out.println(excelLinked.toString()+"-"+updatedExcelLinked.toString());
	           	    		    System.out.println();
        	    		   }
        	    		   else if((Utility.map_notfound).containsKey(listExcelLinked))
        	    		   {
        	    			   	Utility.writeFile("FilesNotFound.txt",absolutePathExtracted+"  -  "+excelLinked.toString(),true);
              	    		  	System.out.println("File not found -  "+excelLinked.toString()+" - in the folder  - "+mainFolder);
        	    		   }
        	    		   else
        	    		   {
	        	    		   for(int i=0;i<relPathExtracted.size();i++)
	        	    		   {
	        	    			   ArrayList<String>listRelPathExtracted=new ArrayList<String>();
	            	    		   ArrayList<String>tempListRelPathExtracted=new ArrayList<String>();
	        	    			   st1 = new StringTokenizer(relPathExtracted.get(i).toString(),"\\//");
	             			       while (st1.hasMoreTokens()) 
	             			       {  
	             			         listRelPathExtracted.add(st1.nextToken());
	             			       }
	             			       for(int j=listRelPathExtracted.size()-size;j<=listRelPathExtracted.size()-1 && j>=0;j++)
	             			       {
	             			    	  tempListRelPathExtracted.add(listRelPathExtracted.get(j));
	             			       }
	             			       flag1=match(listExcelLinked,tempListRelPathExtracted);
	             			       if(flag1)
	             			       {
	             			    	   list1=tempListRelPathExtracted;
	             			    	   	break;
	             			       }
	        	    		   }
	        	    		   if(flag1)
	        	    		   {
	        	    			   boolean flag2=false;
	        	    			   String updatedExcelLinked=excelLinked;
	        	    			   String s=null;
	        	    			   if(!((Utility.getFileExtension(listExcelLinked.get(size-1))).equals("xls")||(Utility.getFileExtension(listExcelLinked.get(size-1))).equals("xlsx")))
            	    			   {
            	    				   s=stripExtension(list1.get(size-1));
            	    				   list1.add(size-1,s);
            	    			   }
	        	    			   for(int i=0;i<size;i++)
	        	    			   {
	        	    				   if(!(listExcelLinked.get(i).equals(list1.get(i))))
	        	    				   {
	        	    					   flag2=true;
	        	    					   updatedExcelLinked=excelLinked.replace(listExcelLinked.get(i), list1.get(i));
	        	    				   }
	        	    			   }
	        	    			   if(flag2)
	        	    			   {
	        	    				  map_modified.put(listExcelLinked,list1);
	        	    				  cell.setCellValue(updatedExcelLinked);
	             	    			  FileOutputStream os=new FileOutputStream(excel);
	             	    			  workbook.write(os);
	             	    			  os.close();
	             	    			  writeFile("LinksModified.txt",absolutePathExtracted+"	 -	"+excelLinked+"	 -	"+updatedExcelLinked.toString()+"\n",true);
		           	    			  System.out.println(excelLinked.toString()+"-"+updatedExcelLinked.toString());
		           	    			  System.out.println();
	        	    			   }
	        	    			   else
	        	    			   {
	        	    				  map_found.put(listExcelLinked,listExcelLinked);
	        	    				  writeFile("CorrectPaths.txt",absolutePathExtracted+"     -     "+excelLinked.toString()+"\n",true);
	             	    			  System.out.println("CorrectPath"+"    -    "+excelLinked);
	             	    			  continue;
	        	    			   }
	        	    		   }
	        	    		   else
	        	    		   {
	        	    			  map_notfound.put(listExcelLinked,listExcelLinked);
	        	    			  Utility.writeFile("FilesNotFound.txt",absolutePathExtracted+"  -  "+excelLinked.toString(),true);
	             	    		  System.out.println("File not found -  "+excelLinked.toString()+" - in the folder  - "+mainFolder);
	        	    		   }		         
	            	          cell=null;
        	    		   }
            	    }
            	    
            	  }
            	}
            workbook.close();
            file.close();
        }
        catch (Exception e)
        {
        	writeFile("Exceptions_file.txt",e+" - \t\t\t - "+absolutePathExtracted,true);
        	System.out.println(e+" - file read - "+absolutePathExtracted);
        	System.out.println(e.getStackTrace());
        }
    }
	public static void readExcel_xls(String mainFolder,String absolutePathExtracted, ArrayList<String>relPathExtracted)
    {
        try
        {
        	File excel=new File(absolutePathExtracted);
            FileInputStream file = new FileInputStream(excel);
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFSheet sheet = workbook.getSheetAt(0);
            if(sheet==null)
            {
            	workbook.close();
            	return;
            }
            Row row = sheet.getRow(19);
            if(row==null)
            {
            	workbook.close();
            	return;
            }
            Iterator<Cell> cellIterator = row.cellIterator();
            int colNo=0;
            boolean flag=false;
            String excelLinked=new String();
            while (cellIterator!=null && cellIterator.hasNext())
            {
                Cell cell = cellIterator.next();
                switch (cell.getCellType())
                {
                    case Cell.CELL_TYPE_STRING:
                    	excelLinked=cell.getStringCellValue();
                        if(excelLinked.contentEquals("EXCEL TO BE LINKED"))
                        {
                        	colNo=(cell.getColumnIndex());
                        	flag=true;
                        	break;
                        }
                    break;
                    case Cell.CELL_TYPE_NUMERIC:
                    	break;
                    default:
                    	break;       
                }
                if(flag)
                	break;
            }
            for (int rowIndex = 20; flag==true && rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            	  row = sheet.getRow(rowIndex);
            	  if (row != null) {
            	    Cell cell = row.getCell(colNo);
            	    if (cell != null && cell.getCellType()==Cell.CELL_TYPE_STRING && row.getCell(colNo).getStringCellValue().length()!=0) 
            	    {
            	    	   boolean flag1=false;
	            	       excelLinked=(row.getCell(colNo).getStringCellValue()).toString();
	            	       StringTokenizer st = new StringTokenizer(excelLinked.toString(),"\\//");  
        	    		   ArrayList<String>listExcelLinked=new ArrayList<String>();
        	    		   while (st.hasMoreTokens()) 
        			       {  
        			         listExcelLinked.add(st.nextToken());
        			       }
        	    		   int size=listExcelLinked.size();
        	    		   if(size==0)
        	    			   continue;
        	    		   StringTokenizer st1=null;
        	    		   ArrayList<String>list1=new ArrayList<String>();
        	    		   if((Utility.map_found).containsKey(listExcelLinked))
        	    		   {
        	    			  writeFile("CorrectPaths.txt",absolutePathExtracted+"     -     "+excelLinked.toString()+"\n",true);
          	    			  System.out.println("CorrectPath"+"    -    "+excelLinked);
          	    			  continue;
        	    		   }
        	    		   else if((Utility.map_modified).containsKey(listExcelLinked))
        	    		   {
        	    			   list1=map_modified.get(listExcelLinked);
        	    			   String updatedExcelLinked=excelLinked;
        	    			   for(int i=0;i<size;i++)
        	    			   {
        	    				   if(!(listExcelLinked.get(i).equals(list1.get(i))))
        	    				   {
        	    					   updatedExcelLinked=excelLinked.replace(listExcelLinked.get(i), list1.get(i));
        	    				   }
        	    			   }
        	    			    cell.setCellValue(updatedExcelLinked);
           	    			    FileOutputStream os=new FileOutputStream(excel);
           	    			    workbook.write(os);
           	    			    os.close();
           	    			  	writeFile("LinksModified.txt",absolutePathExtracted+"	 -	"+excelLinked+"	 -	"+updatedExcelLinked.toString()+"\n",true);
	           	    			System.out.println(excelLinked.toString()+"-"+updatedExcelLinked.toString());
	           	    		    System.out.println();
        	    		   }
        	    		   else if((Utility.map_notfound).containsKey(listExcelLinked))
        	    		   {
        	    			   	Utility.writeFile("FilesNotFound.txt",absolutePathExtracted+"  -  "+excelLinked.toString(),true);
              	    		  	System.out.println("File not found -  "+excelLinked.toString()+" - \t\tin the folder  - "+mainFolder);
        	    		   }
        	    		   else
        	    		   {
	        	    		   for(int i=0;i<relPathExtracted.size();i++)
	        	    		   {
	        	    			   ArrayList<String>listRelPathExtracted=new ArrayList<String>();
	            	    		   ArrayList<String>tempListRelPathExtracted=new ArrayList<String>();
	        	    			   st1 = new StringTokenizer(relPathExtracted.get(i).toString(),"\\//");
	             			       while (st1.hasMoreTokens()) 
	             			       {  
	             			         listRelPathExtracted.add(st1.nextToken());
	             			       }
	             			       for(int j=listRelPathExtracted.size()-size;j<=listRelPathExtracted.size()-1 && j>=0;j++)
	             			       {
	             			    	  tempListRelPathExtracted.add(listRelPathExtracted.get(j));
	             			       }
	             			       flag1=match(listExcelLinked,tempListRelPathExtracted);
	             			       if(flag1)
	             			       {
	             			    	   list1=tempListRelPathExtracted;
	             			    	   	break;
	             			       }
	        	    		   }
	        	    		   if(flag1)
	        	    		   {
	        	    			   boolean flag2=false;
	        	    			   String updatedExcelLinked=excelLinked;
	        	    			   String s=null;
	        	    			   if(!((Utility.getFileExtension(listExcelLinked.get(size-1))).equals("xls")||(Utility.getFileExtension(listExcelLinked.get(size-1))).equals("xlsx")))
            	    			   {
            	    				   s=stripExtension(list1.get(size-1));
            	    				   list1.add(size-1,s);
            	    			   }
	        	    			   for(int i=0;i<size;i++)
	        	    			   {
	        	    				   if(!(listExcelLinked.get(i).equals(list1.get(i))))
	        	    				   {
	        	    					   flag2=true;
	        	    					   updatedExcelLinked=excelLinked.replace(listExcelLinked.get(i), list1.get(i));
	        	    				   }
	        	    			   }
	        	    			   if(flag2)
	        	    			   {
	        	    				  map_modified.put(listExcelLinked,list1);
	        	    				  cell.setCellValue(updatedExcelLinked);
	             	    			  FileOutputStream os=new FileOutputStream(excel);
	             	    			  workbook.write(os);
	             	    			  os.close();
	             	    			  writeFile("LinksModified.txt",absolutePathExtracted+"	 -	"+excelLinked+"	 -	"+updatedExcelLinked.toString()+"\n",true);
		           	    			  System.out.println(excelLinked.toString()+"-"+updatedExcelLinked.toString());
		           	    			  System.out.println();
	        	    			   }
	        	    			   else
	        	    			   {
	        	    				  map_found.put(listExcelLinked,listExcelLinked);
	        	    				  writeFile("CorrectPaths.txt",absolutePathExtracted+"     -     "+excelLinked.toString()+"\n",true);
	             	    			  System.out.println("CorrectPath"+"    -    "+excelLinked);
	             	    			  continue;
	        	    			   }
	        	    		   }
	        	    		   else
	        	    		   {
	        	    			  map_notfound.put(listExcelLinked,listExcelLinked);
	        	    			  Utility.writeFile("FilesNotFound.txt",absolutePathExtracted+"  -  "+excelLinked.toString(),true);
	             	    		  System.out.println("File not found -  "+excelLinked.toString()+" - \t\tin the folder  - "+mainFolder);
	        	    		   }		         
	            	          cell=null;
        	    		   }
            	    }
            	    
            	  }
            	}
            workbook.close();
            file.close();
        }
        catch (Exception e)
        {
        	writeFile("Exceptions_file.txt",e+" - file read - "+absolutePathExtracted,true);
        	System.out.println(e+" - \t\t\t - "+absolutePathExtracted);
        	System.out.println(e.getStackTrace());
        }
    }
	public static void main(String[] args) 
	{  
		try
		{
			map_notfound=new HashMap<ArrayList<String>,ArrayList<String> >();
			   map_found=new HashMap<ArrayList<String>,ArrayList<String> >();
			   map_modified=new HashMap<ArrayList<String>,ArrayList<String> >();
			writeFile("Exceptions_file.txt"," \t\tException Name \t\t\t \t\t\t\t\t\t\t\t\t- file read ",false);
			writeFile("CorrectPaths.txt","\t\tExcel File Read\t\t\t\t-\t\t\t\t\tCorrectPaths (Files Found)\n\n\n",false);
			writeFile("FilesNotFound.txt","\t\tExcel File Read\t\t\t\t-\t\t\t\twrong path written in the given excel file (Files Not Found)\n\n\tEither Spelling Mistakes or File Doesn't Exist\n\n\n",false);
			writeFile("LinksModified.txt","\t\tExcel File Read\t\t\t-\t\t\t\t\t\tOldLink\t\t\t\t-\t\t\t\t\t\tNewLink\n\n\n",false);
			int i;	
		       ArrayList<String>Folders=new ArrayList<String>();
		       System.out.println("Enter the path(ex: D:\\\\DMS\\\\Automation ):");
		       Scanner sc = new Scanner(System.in);
		       String pathInput = sc. nextLine();
		       ArrayList<String>mainFolders=new ArrayList<String>();
		       
		       Folders=Utility.getMainFolders(pathInput);
		       for(i=0;i<Folders.size();i++)
		       {   
		    	   ArrayList<String>arr=new ArrayList<String>();
			       StringTokenizer st = new StringTokenizer(Folders.get(i).toString());  
			       while (st.hasMoreTokens()) 
			       {  
			         arr.add(st.nextToken("\\"));
			       }
			       mainFolders.add(arr.get(arr.size()-1));
		       }
		     
		       for(i=0;i<mainFolders.size() && mainFolders!=null;i++)
		       {
		    	   map_notfound=new HashMap<ArrayList<String>,ArrayList<String> >();
				   map_found=new HashMap<ArrayList<String>,ArrayList<String> >();
				   map_modified=new HashMap<ArrayList<String>,ArrayList<String> >();
		    	   ArrayList<String>path_xlsx=new ArrayList<String>();
			       ArrayList<String>path_xls=new ArrayList<String>();
		    	   getPath(pathInput+"\\"+mainFolders.get(i),path_xlsx,path_xls);
		    	   //System.out.println(path_xls.size());
		    	   for(int j=0;j<path_xlsx.size()&&path_xlsx!=null;j++)
		    	   {
		    		   Utility.readExcel_xlsx(mainFolders.get(i),path_xlsx.get(j).toString(),path_xlsx);
		    	   }
		    	   for(int j=0;j<path_xls.size()&&path_xls!=null;j++)
		    	   {
		    		   Utility.readExcel_xls(mainFolders.get(i),path_xls.get(j).toString(),path_xls);
		    	   }
		       }
		       System.out.println("\n\n\n\t\t\t\t\tProgram Executed Successfully");
		}
		catch(Exception e)
		{
			System.out.println(e.getStackTrace());
		}
		
	 }
}
