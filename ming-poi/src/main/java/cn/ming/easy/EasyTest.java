package cn.ming.easy;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by ming on 2020/11/22.
 */
public class EasyTest {
    String path = "C:\\Users\\ASUS\\Desktop\\huanzi-qch-base-admin-master\\excel\\ming-poi\\";

    private List<DemoData> data(){
        List<DemoData> list = new ArrayList<DemoData>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串"+i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

    //根据list写入Excel
    @Test
    public void simpleWrite() {
        String fileName = path + "EasyTest.xlsx";
        // 这里 需要指定写用哪个class去读，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
        // write (fileName , 格式类)
        // sheet 表名
        // doWrite 写入数据
        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(data());
    }


    @Test
    public void simpleRead() {
        String fileName = path + "EasyTest.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
    }
}
