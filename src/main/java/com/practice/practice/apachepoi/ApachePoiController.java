package com.practice.practice.apachepoi;

import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;


@RestController
public class ApachePoiController {
    @Autowired
    private ApachePoiService apachePoiService;

    @GetMapping("/apache-poi/test")
    public ResponseEntity<Resource> apachePoiTest(HttpServletResponse response) throws IOException{
        final String fileName = "TEST";
        return ResponseEntity
              .ok()
              .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=" + fileName +".xlsx")
              .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
              .body(apachePoiService.test());
    }
    
}
