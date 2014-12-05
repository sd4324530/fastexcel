package com.github.sd4324530.fastexcel;

import com.github.sd4324530.fastexcel.entity.MyTest;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * @author peiyu
 */
public class ExcelTest {

    private static final Logger LOG = LoggerFactory.getLogger(ExcelTest.class);

    @Test
    public void test() {
        readExcel();
    }

    private void readExcel() {
        FastExcel fastExcel = new FastExcel("E:/data.xlsx");
        fastExcel.setSheetName("活动信息数据");
        List<MyTest> tests = fastExcel.parse(MyTest.class);
        if(null != tests && !tests.isEmpty()) {
            for(MyTest myTest : tests) {
                LOG.debug("记录:{}", myTest.toString());
            }
        } else {
            LOG.debug("没有结果");
        }
    }
}
