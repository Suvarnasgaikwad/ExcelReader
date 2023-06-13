package com.smp.ExeclReader;
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.hibernate.Transaction;
import org.hibernate.cfg.Configuration;

public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
        Configuration cfg=new Configuration();
		cfg.configure("Hibernate.cfg.xml");
		SessionFactory factory=cfg.buildSessionFactory();
		Session s=factory.openSession();
		
		Transaction tx=s.beginTransaction();
        try {

            
        	File file =new File("E:\\smp\\Book1.xlsx");
    		FileInputStream fis=new FileInputStream(file);
    		XSSFWorkbook wb = new XSSFWorkbook(fis);
    		XSSFSheet sheet = wb.getSheetAt(0);
    		for(int i=sheet.getFirstRowNum()+1;i<=sheet.getLastRowNum();i++)
    		{
    			Student stud=new Student();
    			Row ro=sheet.getRow(i);
    			for(int j=ro.getFirstCellNum();j<=ro.getLastCellNum();j++)
    			{
    				Cell ce=ro.getCell(j);
    				if(j==0)
    				{
    					stud.setId((int) ce.getNumericCellValue());
    				}
    				if(j==1)
    				{
    					stud.setName(ce.getStringCellValue());
    				}
    				if(j==2)
    				{
    					stud.setAddress(ce.getStringCellValue());
    				}
    				if(j==3)
    				{
    					stud.setSalary(ce.getNumericCellValue());
    				}
    				s.save(stud);
    	    		
    			}
    		}
    		
        }
        catch(Exception e)
    	{
    		e.printStackTrace(); 
    	}
        System.out.println("Data Save Successfully");
		tx.commit();
		s.close();
		factory.close();
   }
}
