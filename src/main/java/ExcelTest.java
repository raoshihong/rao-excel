import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * @author raoshihong
 * @date 11/29/21 9:29 PM
 */
public class ExcelTest {

    public static void main(String[] args) throws Exception{

        ExcelWriter bigWriter = ExcelUtil.getBigWriter();

        Map<String,List<String>> dataGroup = new LinkedHashMap<>();

        // 模拟从数据库中获取数据
        for (int j=0;j<1;j++){
            String date = LocalDate.now().plusDays(j).format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            List<String> dataList = new ArrayList<>();
            for (int i=0;i<100000;i++){
                dataList.add(i+"");
            }
            dataGroup.put(date,dataList);
        }

        // 一new,则BigExcelwriter中的getCurrentRow()又变成0了
        Map<String,BigExcelWriter> bigExcelWriterMap= new HashMap<>();
        // 分页查询
        List<String> data = new ArrayList<>();
        dataGroup.forEach((date, dataList) -> {
            BigExcelWriter bigExcelWriter = bigExcelWriterMap.get(date);
            if (Objects.isNull(bigExcelWriter)) {
                bigWriter.setSheet(date);
                bigExcelWriter = new BigExcelWriter(bigWriter.getWorkbook().getSheet(date));
                bigExcelWriterMap.put(date,bigExcelWriter);
            }


            for (int i=1;i<dataList.size()+1;i++){
                data.add(dataList.get(i-1));
                if (i%200==0) {
//                    bigWriter.setSheet(date);
                    // 这种写法不对,这样会导致同一个sheet在write方法中的getCurrentRow为0,应该是同一个sheet只能有一个Writer,这样行的计数器就是在累加
//                    BigExcelWriter bigExcelWriter = new BigExcelWriter(bigWriter.getWorkbook().getSheet(date));
//                    bigExcelWriterMap.put(date,bigExcelWriter);
                    bigExcelWriter.write(data);
                    data.clear();
                }
            }
        });

        bigWriter.flush(new FileOutputStream(new File("/Users/raoshihong/Desktop/test.xlsx")));

        bigExcelWriterMap.values().forEach(BigExcelWriter::close);
        bigWriter.close();
    }

}
