package com.github.houbb.iexcel.test.write;



import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.core.write.ExcelWriterImpl;
import com.github.houbb.iexcel.metadata.Sheet;
import com.github.houbb.iexcel.test.model.ExcelPropertyIndexModel;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author jipengfei
 * @date 2017/05/31
 */
public class ExcelWriteIndexTest {

    @Test
    public void test1() throws FileNotFoundException {
        OutputStream out = new FileOutputStream("D:\\github\\iexcel\\src\\test\\resources\\1.xlsx");
        try {
            ExcelWriterImpl writer = new ExcelWriterImpl(out, ExcelTypeEnum.XLSX);
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 2, ExcelPropertyIndexModel.class);
            writer.write(getData(), sheet1);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public List<ExcelPropertyIndexModel> getData() {
        List<ExcelPropertyIndexModel> datas = new ArrayList<ExcelPropertyIndexModel>();

        for(int i = 0; i < 1; i++) {
            ExcelPropertyIndexModel model = new ExcelPropertyIndexModel();
//            model.setName("XXXXXXXXXXXXXXXXXXXXXXXXXX");
            model.setAddress("杭州");
            model.setAge("11");
            model.setEmail("7827323@qq.com");
            model.setSax("男");
            model.setHeigh("1123");
            datas.add(model);
        }

        return datas;
    }

}
