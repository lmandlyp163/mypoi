package testExport;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel2007Servlet extends HttpServlet {
	public static final String FILE_SEPARATOR = System.getProperties()
			.getProperty("file.separator");

	@Override
	protected void doPost(HttpServletRequest request,
			HttpServletResponse response) throws ServletException, IOException {
		doGet(request, response);
	}

	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		String docsPath = request.getSession().getServletContext()
				.getRealPath("docs");
		String fileName = "export2007_" + System.currentTimeMillis() + ".xlsx";
		String filePath = docsPath + FILE_SEPARATOR + fileName;
		try {
			// 输出流
			OutputStream os = new FileOutputStream(filePath);
			// 工作区
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFCellStyle style=wb.createCellStyle();
			style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
			style.setBorderRight(HSSFCellStyle.BORDER_SLANTED_DASH_DOT);
			style.setRightBorderColor(HSSFColor.PINK.index);
			style.setBorderBottom(HSSFCellStyle.BORDER_SLANTED_DASH_DOT);
			style.setBottomBorderColor(HSSFColor.PINK.index);
			style.setFillForegroundColor(IndexedColors.RED.getIndex());
//			style.setFillBackgroundColor(IndexedColors.RED.getIndex());
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			XSSFSheet sheet1 = wb.createSheet("test1");
			XSSFSheet sheet2 = wb.createSheet("test2");
			XSSFRow row1_0 = sheet1.createRow(0);
			XSSFRow row2_0 = sheet2.createRow(0);
			row1_0.createCell(0).setCellValue("cell1");
			row1_0.createCell(1).setCellValue("cell2");
			row2_0.createCell(0).setCellValue("cell1");
			row2_0.createCell(1).setCellValue("cell2");
			for (int i = 1; i < 1000; i++) {
				CreationHelper createHelper = wb.getCreationHelper();
				Hyperlink link =createHelper.createHyperlink(Hyperlink.LINK_URL);
				link.setAddress("http://www.baidu.com");
				
				// 创建第一个sheet
				// 生成第一行
				XSSFRow row = sheet1.createRow(i);
				// 给这一行的第一列赋值
				XSSFCell cell0 = row.createCell(0);
				cell0.setCellValue(i);
//				row.createCell(0).setCellType(HSSFColor.PINK.index);
				cell0.setCellStyle(style);
				cell0.setHyperlink(link);
				// 给这一行的第一列赋值
				row.createCell(1).setCellValue(i+1);
				System.out.println(i);
			}
			// 写文件
			wb.write(os);
			// 关闭输出流
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		download(filePath, response);
	}

	private void download(String path, HttpServletResponse response) {
		try {
			// path是指欲下载的文件的路径。
			File file = new File(path);
			// 取得文件名。
			String filename = file.getName();
			// 以流的形式下载文件。
			InputStream fis = new BufferedInputStream(new FileInputStream(path));
			byte[] buffer = new byte[fis.available()];
			fis.read(buffer);
			fis.close();
			// 清空response
			response.reset();
			// 设置response的Header
			response.addHeader("Content-Disposition", "attachment;filename="
					+ new String(filename.getBytes()));
			response.addHeader("Content-Length", "" + file.length());
			OutputStream toClient = new BufferedOutputStream(
					response.getOutputStream());
			response.setContentType("application/vnd.ms-excel;charset=gb2312");
			toClient.write(buffer);
			toClient.flush();
			toClient.close();
		} catch (IOException ex) {
			ex.printStackTrace();
		}
	}
}