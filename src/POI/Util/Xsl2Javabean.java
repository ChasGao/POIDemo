package POI.Util;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import POI.Javabean.Student;

/**
 * 
 * @author 30868
 *
 */
public class Xsl2Javabean {

	/**
	 * 	从xls文件中读取每行的数据，返回对象数组
	 * @return 
	 * @throws IOException
	 */
	public List<Student> readXls() throws IOException {
		List<Student> students = new ArrayList<Student>();
		Student student = null;
		InputStream is = new FileInputStream(
				"D:/!workspaceMyeclipse2015/POIDemo/WebRoot/WEB-INF/student.xls");
		HSSFWorkbook workbook = new HSSFWorkbook(is);
		// 循环读sheet
		for (int numsheet = 0; numsheet < workbook.getNumberOfSheets(); numsheet++) {
			HSSFSheet sheet = workbook.getSheetAt(numsheet);
			if (sheet == null) {
				continue;
			}
			// 循环读每个sheet的每行row;格式：
			// 学号 姓名 年龄 性别 出生日期
			// 10000001 张三 20 男 2015-08-21
			for (int numrow = 1; numrow < sheet.getLastRowNum(); numrow++) {
				HSSFRow row = sheet.getRow(numrow);
				if (row == null) {
					continue;
				}

				student = new Student();
				// 给student的每个属性设值

				HSSFCell cell0 = row.getCell(0);
				if (cell0 == null) {
					student.setId(0);
				} else {
					student.setId(Long.parseLong(getValue(cell0)));
				}

				HSSFCell cell1 = row.getCell(1);
				if (cell1 == null) {
					student.setName("");
				} else {
					student.setName(getValue(cell1));
				}

				HSSFCell cell2 = row.getCell(2);
				if (cell2 == null) {
					student.setAge(0);
				} else {
					student.setAge(Integer.parseInt(getValue(cell2)));
				}

				HSSFCell cell3 = row.getCell(3);
				if (cell3 == null) {
					student.setSex(false);
					;
				} else {
					if ("男".equals(getValue(cell3))) {
						student.setSex(true);
					} else {
						student.setSex(false);
					}
				}

				HSSFCell cell4 = row.getCell(4);
				if (cell4 == null) {
					continue;
				}
				SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
				Date date = null;
				try {
					date = sf.parse(getValue(cell4));
				} catch (ParseException e) {
					e.printStackTrace();
					System.out.println(numrow + "行 birthday转换异常");
				}
				student.setBirthday(date);
				students.add(student);
			}
		}

		return students;
	}

	/**
	 * 得到单元格的值
	 * 
	 * @param HSSFCell
	 *            cell
	 * @return cell单元格的值
	 */
	public String getValue(HSSFCell cell) {
		if (cell.getCellType() == cell.CELL_TYPE_BOOLEAN) {
			// boolean类型
			return String.valueOf(cell.getBooleanCellValue());
		} else if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
			// 数字类型
			return String.valueOf(cell.getNumericCellValue());
		} else {
			// String类型
			return String.valueOf(cell.getStringCellValue());
		}
	}

	public static void main(String[] args) throws Exception {
		SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
		Date date = sf.parse("2015-08-21");
		System.out.println(date);
		boolean b = false;
		System.out.println(b);
		
		List<Student> students=new  Xsl2Javabean().readXls();
		for(Student s:students){
			System.out.println(s.toString());
		}
		
	}
}
