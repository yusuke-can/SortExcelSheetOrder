package com.commandline.SortExcelSheetOrder;

import java.util.Arrays;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@SpringBootApplication
public class SortExcelSheetOrderApplication {

    public static void main(String[] args) {
        SpringApplication.run(SortExcelSheetOrderApplication.class, args);
        Arrays.stream(args).forEach(log::info);
    }


}
