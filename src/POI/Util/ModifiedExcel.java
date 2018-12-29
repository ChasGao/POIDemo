package POI.Util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ModifiedExcel {

	public static void main(String[] args) {
		ModifiedExcel me = new ModifiedExcel();
		me.readExecl("E:\\4-23\\服务请求类型配置数据提取模板-20181123175003.xlsx");
		me.writeExcel("E:\\4-23\\服务请求类型配置数据提取模板-20181123175003.xlsx");
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
	      XSSFSheet sh = workbook.getSheetAt(0);
	      //修改excle表的第5行，从第三列开始的数据
	      for (int i = 2; i < 4; i++) {
	        //对第五行的数据修改
	        sh.getRow(4).getCell((short) i).setCellValue(100210 + i);
	      }
	      //将修改后的文件写出到D:\\excel目录下
	      FileOutputStream os = new FileOutputStream("E:\\4-23\\服务请求类型配置数据提取模板new-20181123175003.xlsx");
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
	            System.out.println("excel表格中取出的数据" + getValue(xCell));
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
