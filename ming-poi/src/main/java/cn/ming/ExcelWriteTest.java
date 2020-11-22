package cn.ming;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;

/**
 * Created by ming on 2020/11/22.
 */
public class ExcelWriteTest {

    String path = "C:\\Users\\ASUS\\Desktop\\huanzi-qch-base-admin-master\\excel\\ming-poi\\";

    @Test
    public void testWrite03() throws Exception {
        //1.创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("Excel表一");
        //3.创建一个行 (1,1)
        Row row = sheet.createRow(0);
        //4.创建一个单元格
        Cell cell = row.createCell(0);
        //5.设置值
        cell.setCellValue("表格(1,1)");
        //(1,2)
        Cell cell1 = row.createCell(1);
        //5.设置值
        cell1.setCellValue(6666);

        Row row1 = sheet.createRow(1);
        Cell cell2 = row1.createCell(0);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell2.setCellValue(time);

        //生成一张表  03版本使用xls结尾
        FileOutputStream stream = new FileOutputStream(path + "03Excel表写入.xls");
        //输出文件
        workbook.write(stream);
        //关闭流
        stream.close();
        System.out.println("文件生成完毕");
    }


    @Test
    public void testWrite07() throws Exception {
        //1.创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("Excel表一");
        //3.创建一个行 (1,1)
        Row row = sheet.createRow(0);
        //4.创建一个单元格
        Cell cell = row.createCell(0);
        //5.设置值
        cell.setCellValue("表格(1,1)");
        //(1,2)
        Cell cell1 = row.createCell(1);
        //5.设置值
        cell1.setCellValue("cell");

        Row row1 = sheet.createRow(1);
        Cell cell2 = row1.createCell(0);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell2.setCellValue(time);
        Cell cell3 = row1.createCell(1);
        cell3.setCellValue("admin");

        //生成一张表  07版本使用xlsx结尾
        FileOutputStream stream = new FileOutputStream(path + "07Excel表写入.xlsx");
        //输出文件
        workbook.write(stream);
        //关闭流
        stream.close();
        System.out.println("文件生成完毕");
    }
}
