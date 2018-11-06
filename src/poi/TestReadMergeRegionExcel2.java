package poi;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

public class TestReadMergeRegionExcel2 {

	
	@Test
	public void testReadExcel() {
	readExcelToObj("C:\\Users\\陈回\\Desktop\\批量导入excel\\标准2\\镀锌钢绞线 5.4-1570Mpa A级.xlsx");
	}
	
	//D:\\poi.xlsx
	
	private void readExcelToObj(String path)
	{
		Workbook wb = null;
		try {
			File file=new File(path);
			wb = WorkbookFactory.create(file);
			//处理sheet的前5行
			readHeadExcel(wb,0,0,0);			
			//处理一个sheet页面的第8行到最后一行
			readExcel(wb, 0, 7, 0);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	
	public void readHeadExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine)
	{
		//获取工作表
		Sheet sheet = wb.getSheetAt(sheetIndex);
		//创建二维数组，用来存储excel里面的信息
		String[][] excels=new String[5][2];
		//将合并的单元格填充到数组
		test2(excels,sheet,startReadLine);
		
		//创建行对象
		Row row = null;
		for (int i = startReadLine; i <5; i++) 
		{
			//获取某一行的信息
			row = sheet.getRow(i);
			//遍历某一行的信息
			for (Cell c : row) 
			{				
				int dex=c.getColumnIndex();
				System.out.println(dex);
				if(dex<=4)
				{
					//判断某一个单元格是不是合并单元格
					boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
					
					if(!isMerge) 
					{
//						c.getRichStringCellValue();
						String value=getCellValue(c);
						excels[i][dex]=value;
						//c.getRichStringCellValue().toString();
					}		
				}
						
			}			
		}
		
		
	}
	
	
	
	//处理一个sheet页面的第8行到最后一行
	private void readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine) 
	{
		
		
		//获取工作表
		Sheet sheet = wb.getSheetAt(sheetIndex);
		//获取最下面一行位置（如果有4行，则为3）
		int lineNum = sheet.getLastRowNum();
		//获得最大的列数（如果最大列有6列则为6）
		int maxColoum=sheet.getRow(0).getPhysicalNumberOfCells();
		for(int i=1;i<lineNum+1;i++)
		{
			if(sheet.getRow(i).getPhysicalNumberOfCells()>maxColoum)
			{
				maxColoum=sheet.getRow(i).getPhysicalNumberOfCells();
			}
		}
		//创建二维数组，用来存储excel里面的信息
		String[][] excels=new String[lineNum+1-startReadLine][maxColoum];
		//将合并的单元格填充到数组
		test1(excels,sheet,startReadLine);	
		//创建行对象
		Row row = null;
		//sheet.getLastRowNum()获取最后一行的行的下标，从0开始计算
		for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) 
		{
			//获取某一行的信息
			row = sheet.getRow(i);
			//遍历某一行的信息
			for (Cell c : row) 
			{				
				int dex=c.getColumnIndex();
				System.out.println(dex);
				//判断某一个单元格是不是合并单元格
				boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
				
				if(!isMerge) 
				{
//					c.getRichStringCellValue();
					String value=getCellValue(c);
					excels[i-startReadLine][dex]=value;
					//c.getRichStringCellValue().toString();
				}				
			}			
		}
		
		
		System.out.println(excels);
		
		
	}
	
	
	  public void test2(String[][] excels,Sheet sheet,int startReadLine)
	  {
		
		  int sheetMergeCount = sheet.getNumMergedRegions();
			for (int i = 0; i < sheetMergeCount; i++) 
			{
				CellRangeAddress range = sheet.getMergedRegion(i);
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				if(lastRow<=4)
				{					
					Row fRow = sheet.getRow(firstRow);        
					Cell fCell = fRow.getCell(firstColumn);  
					String	value=getCellValue(fCell); 					
					for(int j=firstRow;j<firstRow+1;j++)
					{
						for(int m=firstColumn;m<firstColumn+1;m++)
						{
							excels[j][m]=value;
						}
					}			
				}
				
				System.out.println(excels);
																					
			}  
		  
		  
	  }
	
	
	
	
 
	  public void test1(String[][] excels,Sheet sheet,int startReadLine)
	  {
		  int sheetMergeCount = sheet.getNumMergedRegions(); 
			for (int i = 0; i < sheetMergeCount; i++) 
			{
				CellRangeAddress range = sheet.getMergedRegion(i);
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				if(firstRow>=startReadLine)
				{
					Row fRow = sheet.getRow(firstRow);        
					Cell fCell = fRow.getCell(firstColumn);  
					String	value=getCellValue(fCell); 
					
					for(int j=firstRow;j<lastRow+1;j++)
					{
						for(int m=firstColumn;m<lastColumn+1;m++)
						{
							excels[j-startReadLine][m]=value;
						}
					}			
				}
																					
			}  
	  }
	
	
	
	
	
	
	
	
	private boolean isMergedRegion(Sheet sheet,int row ,int column) 
	{  
		int sheetMergeCount = sheet.getNumMergedRegions(); 
		for (int i = 0; i < sheetMergeCount; i++) 
		{
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if(row >= firstRow && row <= lastRow)
			{
				if(column >= firstColumn && column <= lastColumn)
				{
					return true;
				}
			}  
		}  
		return false;
	}
	
	public String getCellValue(Cell cell)
	{   
		if(cell == null)
		{
			return "";  
		}
			 
		if(cell.getCellType() == Cell.CELL_TYPE_STRING)
		{          
			return cell.getStringCellValue();    
		}
		else if(cell.getCellType() == Cell.CELL_TYPE_BOOLEAN)
		{             
			return String.valueOf(cell.getBooleanCellValue());  
		}
		else if(cell.getCellType() == Cell.CELL_TYPE_FORMULA)
		{    
			return cell.getCellFormula() ;          
		}
		else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
		{                 
			return String.valueOf(cell.getNumericCellValue());   
		}    
		return ""; 
	}  
	
//	public String getMergedRegionValue(Sheet sheet ,int row , int column)
//	{      
//		//得到合并单元格的个数
//		int sheetMergeCount = sheet.getNumMergedRegions();  
//		//遍历合并单元格
//		for(int i = 0 ; i < sheetMergeCount ; i++)
//		{         
//			//得到第i个合并单元格
//			CellRangeAddress ca = sheet.getMergedRegion(i);   
//			//得到某个合并单元格的开始行，结束行，开始列，结束列
//			int firstColumn = ca.getFirstColumn();      
//			int lastColumn = ca.getLastColumn();      
//			int firstRow = ca.getFirstRow();   
//			int lastRow = ca.getLastRow();       
//			if(row >= firstRow && row <= lastRow)
//			{                       
//				if(column >= firstColumn && column <= lastColumn)
//				{               
//					//如果得到的是合并的单元格，得到这个合并单元格
//					Row fRow = sheet.getRow(firstRow);        
//					Cell fCell = fRow.getCell(firstColumn);  
//					//获得单元格数据
//					return getCellValue(fCell) ;         
//				}      
//			}      
//		}        
//		return null ; 
//	}  
//	
	
	
	
	
	
	
	

	
	
	
	
	
	
	
	
	
}
