package com.excelRead.poi.controller;

import com.excelRead.poi.service.ExcelDataService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.FileNotFoundException;
import java.io.IOException;

@RestController
@RequestMapping("/hello")
public class HelloWorldCont {

    @GetMapping("/world")
    public static void helloWorld(){
        System.out.println("hello World");
    }
    @Autowired
    private ExcelDataService excelDataService;

    @PostMapping("/hi")
    public ResponseEntity<String> processExcelDataRequest(@RequestParam("seriesFilePath") String seriesFilePath, @RequestParam("seasonFilePath") String seasonFilePath, @RequestParam("episodeFilePath") String episodeFilePath){
        try {
            excelDataService.processExcelFiles(seriesFilePath, seasonFilePath, episodeFilePath);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return ResponseEntity.status(HttpStatus.OK).body("Excel files processed successfully");
    }
}
