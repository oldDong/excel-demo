package com.dongzj.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.InputStream;
import java.net.URLEncoder;

/**
 * Excel导入和导出
 * User: dongzj
 * Mail: dongzj@shinemo.com
 * Date: 2018/11/19
 * Time: 11:34
 */
public class ExcelUtil {

    /**
     * 导入本地文件到mysql
     *
     * @param path
     */
    public static void importExcel(String path) {
        if (StringUtils.isBlank(path)) {
            return;
        }
        String fileType = path.substring(path.lastIndexOf(".") + 1);
        InputStream is = null;
        try {
            is = new FileInputStream(path);
            //获取工作簿
            Workbook wb = null;
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook(is);
            } else if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook(is);
            } else {
                return;
            }

            //读取第一个工作页sheet
            Sheet sheet = wb.getSheetAt(0);
            //第一行为标题
            for (Row row : sheet) {
                for (Cell cell : row) {
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    String value = cell.getStringCellValue();
                    System.out.println(value);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 导出数据到excel
     */
    public static void exportExcel(HttpServletResponse response) {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        Sheet sheet = xssfWorkbook.createSheet();
        //创建标题行
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("企业名称");
        row.createCell(1).setCellValue("员工名称");
        row.createCell(2).setCellValue("员工标签");
        for (int i = 0; i < 3; i++) {
            row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue("达达科技");
            row.createCell(1).setCellValue("张三");
            row.createCell(2).setCellValue("快门手");
        }
        String name = "test1159";
        exportExcel(response, name, xssfWorkbook);
    }

    /**
     * 导出EXCEL Java
     *
     * @param response
     * @param fileName
     * @param xssfWorkbook
     */
    private static void exportExcel(HttpServletResponse response, String fileName, XSSFWorkbook xssfWorkbook) {
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment;filename*=utf-8'zh_cn'" + URLEncoder.encode(fileName, "UTF-8"));
            ServletOutputStream out = response.getOutputStream();
            xssfWorkbook.write(out);
            System.out.println("export " + fileName + "success");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            xssfWorkbook = null;
        }
    }

    public static void main(String[] args) {
        String path = "/Users/dongzj/Workspaces/test/excel-demo/src/main/resources/test.xlsx";
        importExcel(path);
    }
}
