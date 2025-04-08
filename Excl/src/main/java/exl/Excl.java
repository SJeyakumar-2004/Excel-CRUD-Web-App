package exl;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Excl
{
	public static void main(String[] args) 
	{
		Workbook wb=new XSSFWorkbook();
		Sheet sheet=wb.createSheet();
		
		Object[] [] data= {
				{"Name","Age","City"},
				{"jk",21,"mylai"},
				{"indhu",21,"mdv"}
		};
		
		int rowNum=0;
		
		for(Object[] rowData:data)
		{
			Row row=sheet.createRow(rowNum++);
			int cellNum=0;
			for(Object cellData:rowData)
			{
				Cell cell=row.createCell(cellNum++);
				if(cellData instanceof String)
				{
					cell.setCellValue((String) cellData);
				}
				else if(cellData instanceof Integer)
				{
					cell.setCellValue((Integer) cellData);
				}
			}
		}
		
		try(FileOutputStream file=new FileOutputStream("c:/Excel/ex.xlsx"))
		{
			wb.write(file);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		finally
		{
			try
			{
				wb.close();
			}
			catch(Exception e)
			{
				e.printStackTrace();
			}
		}
	}
}













// EXCEL FILE READ
/*
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excl {

	public static void main(String[] args) throws IOException
	{
		String loc="c:/Excel/Excel.xlsx";
		
		FileInputStream fil=new FileInputStream(loc);
		
		XSSFWorkbook wb=new XSSFWorkbook(fil);
		
		XSSFSheet sheet=wb.getSheetAt(0);
		
		int rowcount=sheet.getLastRowNum();
		int colcount=sheet.getRow(1).getLastCellNum();
		
		for(int i=0;i<=rowcount;i++)
		{
			XSSFRow row=sheet.getRow(i);
		{
			for(int j=0;j<=colcount;j++)
			{
				XSSFCell cell=row.getCell(j);
			
			System.out.println(cell);
			}
		}
		}

	}

}     */


                         // EXCEL FILE READ ANOTHER METHOD
/*
 
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excl {

	public static void main(String[] args) 
	{
		String loc="c:/Excel/Excel.xlsx";
		try(
		FileInputStream fil=new FileInputStream(loc);
		
		XSSFWorkbook wb=new XSSFWorkbook(fil)){
		
		XSSFSheet sheet=wb.getSheetAt(0);
		
		for(Row row:sheet)
		{
			for(Cell cell:row)
			{
				switch(cell.getCellType())
				{
				case STRING:
					System.out.print(cell.getStringCellValue()+ "\t");
					break;
					
				case NUMERIC:
					System.out.print(cell.getNumericCellValue()+ "\t");
					break;
					
				default:
					System.out.println("unknown \t");
					
				}
			}
			System.out.println();
		}
	}
	catch(IOException e) {
		e.printStackTrace();
	}
}
}

*/


                             // EXCEL FILE DB WRITE
/*
import java.io.FileInputStream;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Excl {

	public static void main(String[] args)
	{
		String loc="c:/Excel/Excel.xlsx";
		
		try(
				FileInputStream fil=new FileInputStream(loc);
				XSSFWorkbook wb=new XSSFWorkbook(fil);
		        Connection c=DriverManager.getConnection("jdbc:mysql://localhost:3306/emp","root","root"))
		        
				{
					XSSFSheet sheet=wb.getSheetAt(0);
					int rowcount=sheet.getLastRowNum();
					int colcount=sheet.getRow(0).getLastCellNum();
				
					String sql = "insert into user(Name, Age, City) values(?, ?, ?)";

					try(PreparedStatement ps=c.prepareStatement(sql))
					
							{
								for(int i=1;i<=rowcount;i++)
								{
									XSSFRow row=sheet.getRow(i);
									if(row!=null)
									{
										for(int j=0;j<colcount;j++)
										{
											XSSFCell cell=row.getCell(j);
											String value=(cell!=null)? cell.toString():null;
										    ps.setString(j+1, value);
										    
										}
										
										ps.executeUpdate();
									}
								}
								System.out.println("Data inserted successfully");
							}
				}
		catch(Exception e)
		{
			e.printStackTrace();
		}
			
	}
}
*/	

                             // EXCEL FILE DB WRITE NOT ALLOW DUPLICATE

/*
import java.io.FileInputStream;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excl {

	public static void main(String[] args) 
	{
		String loc="c:/Excel/Excel.xlsx";
		
		try(
				FileInputStream fil=new FileInputStream(loc);
				XSSFWorkbook wb=new XSSFWorkbook(fil);
				Connection c=DriverManager.getConnection("jdbc:mysql://localhost:3306/emp","root","root"))
		{
			XSSFSheet sheet=wb.getSheetAt(0);
			int rowcount=sheet.getLastRowNum();
			int colcount=sheet.getRow(0).getLastCellNum();
			String sql1="insert into user(Name,Age,City) values(?,?,?)";
			String sql2="select count(*) from user where name=?";
			try(
					PreparedStatement ps1=c.prepareStatement(sql1);
					PreparedStatement ps2=c.prepareStatement(sql2))
			{
				for(int i=1;i<=rowcount;i++)
				{
					XSSFRow row=sheet.getRow(i);
					if(row!=null)
					{
						String name=null;
						for(int j=0;j<colcount;j++)
						{
							XSSFCell cell=row.getCell(j);
							String value=(cell!=null)? cell.toString():null;
							if(j==0)
							{
								name=value;
							}
							ps1.setString(j + 1, value);
							
						}
						ps2.setString(1,name);
						ResultSet rs=ps2.executeQuery();
						if(rs.next() && rs.getInt(1)==0)
						{
							ps1.executeUpdate();
							
						}
						else
						{
							System.out.println("remove duplicate name:"+name);
						}
					}
				}
				System.out.println("Data Inserted");
				
			}
			}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
}
	*/


                //CREATE NEW EXCEL FILE AND WRITE


/*
 import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
class Excl
{
	public static void main(String[] args)
	{
		Workbook wb=new XSSFWorkbook();
		Sheet sheet=wb.createSheet();
		
		Object[][] data= {
				{"Name","Age","City"},
				{"Jk",21,"Mylai"},
				{"indhu",21,"MDV"},
				{"nila",21,"MDV"}
		};
		
		int rowNum=0;
		for(Object[] rowData:data)
		{
			Row row=sheet.createRow(rowNum++);
			int cellNum=0;
			for(Object cellData:rowData)
			{
				Cell cell=row.createCell(cellNum++);
				if(cellData instanceof String)
				{
					cell.setCellValue((String) cellData);
				}
				else if(cellData instanceof Integer)
				{
					cell.setCellValue((Integer) cellData);
				}
			}
		}
		
		try(
				FileOutputStream fil=new FileOutputStream(new File("c:/Excel/Example.txt")))
		{
			wb.write(fil);
			System.out.println("Successfully");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		finally {
		try {
		  wb.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		}	
	}
}
	*/



        // NORMAL TXT FILE READ

/*
import java.io.BufferedReader;
import java.io.FileReader;

class Excl
{
	public static void main(String[] args) 
	{
		String file="c:/Excel/jk.txt";
		try(BufferedReader br=new BufferedReader(new FileReader(file)))
		{
			String line;
			while((line=br.readLine())!=null)
			{
				System.out.println(line);
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
}
*/



           // NORMAL TXT FILE WRITER

/*
import java.io.BufferedWriter;
import java.io.FileWriter;
class Excl
{
	public static void main(String[] args) 
	{
		String fil="c:/Excel/indhu.txt";
		try(BufferedWriter bw=new BufferedWriter(new FileWriter(fil)))
		{
			bw.write("JEYAKUMAR");
			bw.newLine();
			bw.write("Indhumathi");
			System.out.println("Success");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
	}
}
*/










