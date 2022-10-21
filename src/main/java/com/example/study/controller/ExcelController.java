package com.example.study.controller;

import com.example.study.common.ExcelUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.List;

@Controller
@RequestMapping("/excel")
public class ExcelController {

    @Autowired
    private com.example.study.service.ExcelService ExcelService;
    @RequestMapping("/selectExcel")
    public String selectExcel(){
        return "excel";
    }


    /*
     * 엑셀 첫번째시트를 불러와 읽는다
     */
    @RequestMapping("/readExcel")
    public String readExcel(@RequestParam("file") MultipartFile file, Model model)
            throws Exception {
        model.addAttribute("dataList", ExcelService.searchExcelList(file));
        return "excelList";
    }
}
