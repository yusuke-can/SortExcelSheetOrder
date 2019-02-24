package com.commandline.SortExcelSheetOrder;

import java.nio.file.Path;
import java.nio.file.Paths;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@RunWith(SpringRunner.class)
@SpringBootTest
public class SortExcelSheetOrderApplicationTests {

    private Path configYmlPath = Paths.get(System.getProperty("user.dir"),
            "src/test/java",
            this.getClass().getPackage().getName().replace('.', '\\'),
            "SortExcelSheetOrderConfigTest.yml");

    @Test
    public void contextLoads() {

        SortExcelSheetOrderApplication.main(new String[] {
                configYmlPath.toString()
                });
    }

}
