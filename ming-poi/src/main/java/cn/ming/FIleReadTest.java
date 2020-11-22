package cn.ming;

import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.Date;

/**
 * Created by ming on 2020/11/22.
 */
public class FIleReadTest {

    String path = "C:\\Users\\ASUS\\Desktop\\huanzi-qch-base-admin-master\\excel\\ming-poi\\";

    @Test
    public void testRead03() throws Exception{
        //获取工作流
        FileInputStream inputStream = new FileInputStream(path+"03Excel表写入.xls");
        //1.创建工作簿
        Workbook workbook = new HSSFWorkbook(inputStream);
        //2.得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3.得到行
        Row row = sheet.getRow(0);
        //4.得到列
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue());

        Cell cell1 = row.getCell(1);
        System.out.println(cell1.getNumericCellValue());
        //关闭流
        inputStream.close();
    }


    @Test
    public void testRead07() throws Exception{
        //获取工作流
        FileInputStream inputStream = new FileInputStream(path+"07Excel表写入.xlsx");
        //1.创建工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);
        //2.得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3.得到行
        Row row = sheet.getRow(0);
        //4.得到列
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue());

//        Cell cell1 = row.getCell(1);
//        System.out.println(cell1.getNumericCellValue());
        //关闭流
        inputStream.close();
    }

    //读取不同类型
    @Test
    public void testCellType() throws Exception{
        //获取文件流
        FileInputStream inputStream = new FileInputStream(path+"07Excel表写入.xlsx");
        //1.创建工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        //2.获取标题内容
        Row rowTitle = sheet.getRow(0);
        if(rowTitle != null){
            //获取所有列数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null){
                    CellType cellType = cell.getCellTypeEnum();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        //3.获取表的内容
        //读取行
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if(rowData != null){
                //读取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("【"+(rowNum+1)+"-"+(cellNum+1)+"】"); //打印行标列标
                    Cell cell = rowData.getCell(cellNum);
                    //匹配列的数据类型
                    if (cell != null){
                        CellType cellType = cell.getCellTypeEnum();
                        String cellValue = "";

                        switch (cellType){
                            case STRING:
                                System.out.print("[String]"); //字符串
                                cellValue= cell.getStringCellValue();
                                break;
                            case BOOLEAN:
                                System.out.print("[Boolean]"); //布尔
                                cellValue= String.valueOf(cell.getBooleanCellValue());
                                break;
                            case BLANK:
                                System.out.print("[Blank]");  //空
                                break;
                            case NUMERIC:
                                System.out.print("[Bumeric]"); //数字(日期、普通数字)
                                if (HSSFDateUtil.isCellDateFormatted(cell)){
                                    System.out.println("[日期]");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    System.out.print("[转换为字符串]");
                                    //如果不是日期格式，防止数字过长
                                    HSSFDataFormatter hssfDataFormat = new HSSFDataFormatter();
                                    cellValue = hssfDataFormat.formatCellValue(cell);
                                }
                                break;
                            case ERROR:
                                System.out.print("[数据类型错误]"); //错误
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        inputStream.close();
    }


    @Test
    public void testFormula() throws Exception{
        //获取文件流
        FileInputStream inputStream = new FileInputStream(path+"07Excel表写入.xlsx");
        //1.创建工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);

        //获取计算公司
        XSSFFormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        //输出单元格内容
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType){
            case FORMULA:   //公式
                String formula = cell.getCellFormula();
                System.out.println(formula);

                //计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }
    }
}
