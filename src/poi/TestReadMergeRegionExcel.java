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


public class TestReadMergeRegionExcel {
	
	@Test
	public void testReadExcel() {
	readExcelToObj("d:\\poi.xlsx");
	}
	
	
	
	/*** 读取excel数据*
	 * 
	 *  @param path
	 */
//	
	private void readExcelToObj(String path)
	{
		Workbook wb = null;
		try {
			File file=new File(path);
			wb = WorkbookFactory.create(file);
			readExcel(wb, 0, 0, 0);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	
	
	
	
	
	/*** 读取excel文件
	 *  @param wb 
	 * @param sheetIndex sheet页下标：从0开始
	 * @param startReadLine 开始读取的行:从0开始 
	 * @param tailLine 去除最后读取的行
	 */
	//
	private void readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine) 
	{
		//获取工作表
		Sheet sheet = wb.getSheetAt(sheetIndex);
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
				//判断某一个单元格是不是合并单元格
				boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
				
				if(isMerge) 
				{
					//是则去合并单元格获取单元里面的信息
					String rs =getMergedRegionValue(sheet, row.getRowNum(),c.getColumnIndex());
					System.out.print(rs + "");
				 }
				else
				 {
					//不是则直接获取单元格信息
					System.out.print(c.getRichStringCellValue()+"");
				 }
				}
			System.out.println();
			}
		}


	
	






    //
	public String getMergedRegionValue(Sheet sheet ,int row , int column)
	{      
		//得到合并单元格的个数
		int sheetMergeCount = sheet.getNumMergedRegions();  
		//遍历合并单元格
		for(int i = 0 ; i < sheetMergeCount ; i++)
		{         
			//得到第i个合并单元格
			CellRangeAddress ca = sheet.getMergedRegion(i);   
			//得到某个合并单元格的开始行，结束行，开始列，结束列
			int firstColumn = ca.getFirstColumn();      
			int lastColumn = ca.getLastColumn();      
			int firstRow = ca.getFirstRow();   
			int lastRow = ca.getLastRow();       
			if(row >= firstRow && row <= lastRow)
			{                       
				if(column >= firstColumn && column <= lastColumn)
				{               
					//如果得到的是合并的单元格，得到这个合并单元格
					Row fRow = sheet.getRow(firstRow);        
					Cell fCell = fRow.getCell(firstColumn);  
					//获得单元格数据
					return getCellValue(fCell) ;         
				}      
			}      
		}        
		return null ; 
	}  
	



	
	
	
	
	





	
	
//private boolean isMergedRow(Sheet sheet,int row ,int column) 
//{ 
//	int sheetMergeCount = sheet.getNumMergedRegions(); 
//	for (int i = 0; i < sheetMergeCount; i++)
//	{
//		CellRangeAddress range = sheet.getMergedRegion(i);
//		int firstColumn = range.getFirstColumn();
//		int lastColumn = range.getLastColumn();
//		int firstRow = range.getFirstRow();
//		int lastRow = range.getLastRow();
//		if(row == firstRow && row == lastRow)
//		{
//			if(column >= firstColumn && column <= lastColumn)
//			{
//				return true;
//			}
//		}  
//	}  
//	return false;
//}

	
	


	







//判断是否是合并单元格
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






//private boolean hasMerged(Sheet sheet) 
//{
//
//    return sheet.getNumMergedRegions() > 0 ? true : false;
//
//}

	








//private void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
//{
//	sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
//} 




/** * 获取单元格的值 * @param cell * @return */ 
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





	
}