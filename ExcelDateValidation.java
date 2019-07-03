package POI.Util;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDateValidation {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
//		dateValidate1();
		dateValidate2();
	}


	/**
	 * Check the value a user enters into a cell against one or more predefined value(s).
	 * 校验用户输入的单元格值是否为	 预定义值
	 */
	public static void dateValidate1() {
		
		 XSSFWorkbook workbook = new XSSFWorkbook();
		  XSSFSheet sheet = workbook.createSheet("Data Validation");
		  XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
		  XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
		    dvHelper.createExplicitListConstraint(new String[]{"11", "21", "31"});
		  CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
		  XSSFDataValidation validation =(XSSFDataValidation)dvHelper.createValidation(
		    dvConstraint, addressList);

		  // Here the boolean value false is passed to the setSuppressDropDownArrow()
		  // method. In the hssf.usermodel examples above, the value passed to this
		  // method is true.            
		  validation.setSuppressDropDownArrow(false);

		  // Note this extra method call. If this method call is omitted, or if the
		  // boolean value false is passed, then Excel will not validate the value the
		  // user enters into the cell.
		  validation.setShowErrorBox(true);
		  validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
		  validation.createErrorBox("title", "Message Text");
		  
		  XSSFDataValidationConstraint dvConstraint1 = (XSSFDataValidationConstraint)
				    dvHelper.createExplicitListConstraint(new String[]{"22", "32", "42"});
		  CellRangeAddressList addressList1 = new CellRangeAddressList(0, 0, 1, 1);
		  XSSFDataValidation validation1 =(XSSFDataValidation)dvHelper.createValidation(
				    dvConstraint1, addressList1);
		  validation1.setSuppressDropDownArrow(false);
		  validation1.setShowErrorBox(true);
		  
		  sheet.addValidationData(validation);
		  sheet.addValidationData(validation1);


		  FileOutputStream os;
		try {
			os = new FileOutputStream("E:\\GitRepository\\POIDemo\\doc\\表格验证.xlsx");
			os.flush();
			workbook.write(os);			
		} catch (IOException e) {
			e.printStackTrace();
		}		
	}

	/**
	 * Drop Down Lists
	 * 下拉列表
	 */
	public static void dateValidate2() {
		 XSSFWorkbook workbook = new XSSFWorkbook();
		  XSSFSheet sheet = workbook.createSheet("Data Validation");
		  XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
		  XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
		    dvHelper.createExplicitListConstraint(new String[]{"11", "21", "31"});
		  CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
		  XSSFDataValidation validation = (XSSFDataValidation)dvHelper.createValidation(
		    dvConstraint, addressList);
		  validation.setShowErrorBox(true);
		  validation.setSuppressDropDownArrow(true);
		  
		  XSSFDataValidationConstraint dvConstraint1 = (XSSFDataValidationConstraint)
				    dvHelper.createExplicitListConstraint(new String[]{"22", "32", "42"});
				  CellRangeAddressList addressList1 = new CellRangeAddressList(0, 10, 1, 1);
				  XSSFDataValidation validation1 = (XSSFDataValidation)dvHelper.createValidation(
				    dvConstraint1, addressList1);
				  validation1.setShowErrorBox(true);
				  validation1.setSuppressDropDownArrow(true);
			  //Messages On Error: 同dateValidate1
				  
		  sheet.addValidationData(validation);
		  sheet.addValidationData(validation1);


		  FileOutputStream os;
		try {
			os = new FileOutputStream("E:\\GitRepository\\POIDemo\\doc\\表格验证2.xlsx");
			os.flush();
			workbook.write(os);			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}


}
