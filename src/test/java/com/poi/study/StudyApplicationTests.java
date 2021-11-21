package com.poi.study;

import com.poi.study.service.PoiService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class StudyApplicationTests {

    @Autowired
    private PoiService poiService;

    @Test
    void contextLoads() {

    }

    @Test
    public void test1() throws Exception{
        poiService.test();
    }

}
