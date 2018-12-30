package POI.Util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ModifiedExcel {

	public static void main(String[] args) {
		ModifiedExcel me = new ModifiedExcel();
//		me.readExecl("G:\\GitRepository\\POIDemo\\doc\\提取模板.xlsx");
//		me.writeExcel("G:\\GitRepository\\POIDemo\\doc\\提取模板.xlsx");
		me.modifiedExcel();
	}

	public void modifiedExcel(){
		//Sept 1
		List<Map<String, String>> targetOrgAndSRTypeIdList = new ArrayList<>();
		
		try{
		String targetOrgPath = "G:\\GitRepository\\POIDemo\\doc\\二线目标机构.xlsx";
		
		XSSFWorkbook xwb = new XSSFWorkbook(targetOrgPath);
		XSSFSheet sheet = xwb.getSheet("Sheet1");
		XSSFRow xRow = null;
		
		System.out.println("二线目标机构文档行数: "+ sheet.getLastRowNum());
		for (int numRow = 0; numRow <= sheet.getLastRowNum(); numRow++) {
    	  xRow = sheet.getRow(numRow);
          if (xRow == null || numRow ==0) {
            continue;
          }	     
          //TODO 待续 
          Map<String, String> targetOrgAndSRTypeIdMap = new HashMap<>();
          String SRTypeId = xRow.getCell(0).getStringCellValue();
          
          String targetOrg = xRow.getCell(2).getStringCellValue();
          targetOrgAndSRTypeIdMap.put("srTypeId", SRTypeId);   
          targetOrgAndSRTypeIdMap.put("targetOrg", targetOrg);
          targetOrgAndSRTypeIdList.add(targetOrgAndSRTypeIdMap);
          
		}  
		System.out.println("targetOrgAndSRTypeIdList大小是: " + targetOrgAndSRTypeIdList.size());

		}catch(Exception e){
			e.printStackTrace();	  
	    }
		
		// Sept 2
		
	    try {
	    	 String templatePath = "G:\\GitRepository\\POIDemo\\doc\\提取模板.xlsx";
		      //传入的文件
		      FileInputStream fileInput = new FileInputStream(templatePath);
		      //poi包下的类读取excel文件
		      // 创建一个webbook，对应一个Excel文件
		      XSSFWorkbook workbook = new XSSFWorkbook(fileInput);
		      //对应Excel文件中的sheet 
		      XSSFSheet sh = workbook.getSheet("模板");
		      XSSFRow row = null;
		      
		      XSSFCell cellSRTypeId = null;
		      XSSFCell cellTargetOrg = null;
		      String cellValueSRTypeId =null;
//		      String cellValueTargetOrg =null;
		      
		      for(int numRow = 0; numRow <= sh.getLastRowNum(); numRow++){
		    	  row = sh.getRow(numRow);
		    	  if(row == null || numRow == 0){
		    		  continue;
		    	  }
		    	  cellSRTypeId = row.getCell(1);
		    	  cellTargetOrg = row.getCell(8);
		    	  cellValueSRTypeId = cellSRTypeId.getStringCellValue();
//		    	  cellValueTargetOrg = cellTargetOrg.getStringCellValue();
		    	  
		    	  if(cellValueSRTypeId == null || cellValueSRTypeId.equals(""))
		    		  continue;
		    	  
		    	  String srTypeId = null;
		    	  String targetOrg = null;
		    	  for(Map<String, String>  targetOrgAndSRTypeId:targetOrgAndSRTypeIdList){
		    		  
		    		  srTypeId =  targetOrgAndSRTypeId.get("srTypeId");
		    		  targetOrg =  targetOrgAndSRTypeId.get("targetOrg");
		    		  if(srTypeId == null || srTypeId.equals("") || targetOrg == null || targetOrg.equals("") )
		    			  System.out.println("targetOrgAndSRTypeIdList 中的 Map 有空值！！！");
		    		  
		    		  if(cellValueSRTypeId.equals(srTypeId)){
		    			  cellTargetOrg.setCellValue(targetOrg);
		    			  
		    		  }
		    		  
		    	  }
		    	  
		      }
		      
		      FileOutputStream os = new FileOutputStream("G:\\GitRepository\\POIDemo\\doc\\提取模板 - 副本.xlsx");
		      os.flush();
		      //将Excel写出
		      workbook.write(os);
		      //关闭流
		      fileInput.close();
		      os.close();
		    } catch (IOException e) {
		      e.printStackTrace();
		    }
		
		
		
	}
	//修改excel表格，path为excel修改前路径（E:\\4-23\\服务请求类型配置数据提取模板new-20181123175003.xlsx）
	  public void writeExcel(String path) {
	    try {
	      //传入的文件
	      FileInputStream fileInput = new FileInputStream(path);
	      //poi包下的类读取excel文件

	      // 创建一个webbook，对应一个Excel文件
	      XSSFWorkbook workbook = new XSSFWorkbook(fileInput);
	      //对应Excel文件中的sheet，0代表第一个
//	      XSSFSheet sh = workbook.getSheetAt(0);
	      XSSFSheet sh = workbook.getSheet("模板");
	      XSSFCell cell = sh.getRow(2).getCell(8);
	      cell.setCellValue("1000");
	      //修改excle表的第5行，从第2列开始的数据
/*	      for (int i = 1; i < 4; i++) {
	        //对第3行的数据修改
	        sh.getRow(2).getCell(i).setCellValue(100210 + i);
	      }*/
	      //将修改后的文件写出到D:\\excel目录下
	      FileOutputStream os = new FileOutputStream("D:\\!workspaceeclipse\\POIDemo\\doc\\提取模板 - 副本.xlsx");
	      // FileOutputStream os = new FileOutputStream("D:\\test.xlsx");//此路径也可写修改前的路径，相当于在原来excel文档上修改
	      os.flush();
	      //将Excel写出
	      workbook.write(os);
	      //关闭流
	      fileInput.close();
	      os.close();
	    } catch (IOException e) {
	      e.printStackTrace();
	    }
	  }
	  
	//读取excel表格中的数据，path代表excel路径
	  public void readExecl(String path) {
	    try {
//	    	path = "G:\\GitRepository\\POIDemo\\WebRoot\\WEB-INF\\book2.xlsx";
	    	File file = new File(path);
	    	boolean exist = file.exists();
	    	System.out.println(path + ", exists: " + exist);
	      //读取的时候可以使用流，也可以直接使用文件名
	      XSSFWorkbook xwb = new XSSFWorkbook(path);
	      //循环工作表sheet
	      for (int numSheet = 0; numSheet < xwb.getNumberOfSheets(); numSheet++) {
	        XSSFSheet xSheet = xwb.getSheetAt(numSheet);
	        if (xSheet == null) {
	          continue;
	        }
	        //循环行row
	        for (int numRow = 0; numRow <= xSheet.getLastRowNum(); numRow++) {
	          XSSFRow xRow = xSheet.getRow(numRow);
	          if (xRow == null) {
	            continue;
	          }
	          //循环列cell
	          for (int numCell = 0; numCell <= xRow.getLastCellNum(); numCell++) {
	            XSSFCell xCell = xRow.getCell(numCell);
	            if (xCell == null) {
	              continue;
	            }
	            //输出值
	            System.out.println("excel表格中取出的数据		" + getValue(xCell));
	          }
	        }

	      }

	    } catch (IOException e) {
	      e.printStackTrace();
	    }
	  }

	  /**
	   * 取出每列的值
	   *
	   * @param xCell 列
	   * @return
	   */
	  private String getValue(XSSFCell xCell) {
	    if (xCell.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN) {
	      return String.valueOf(xCell.getBooleanCellValue());
	    } else if (xCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
	      return String.valueOf(xCell.getNumericCellValue());
	    } else {
	      return String.valueOf(xCell.getStringCellValue());
	    }
	  }
}
