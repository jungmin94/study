package com.example.study.controller;

import com.example.study.common.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Controller
@RequestMapping("/excel")
public class ExcelController {

    @RequestMapping("/selectExcel")
    public String selectExcel(){
        return "excel";
    }


    /*
     * 엑셀 첫번째시트를 불러와 읽는다
     */
    @RequestMapping("/readExcel")
    public String readExcel(@RequestParam("file") MultipartFile file, Model model)
            throws IOException {
        //getWorkbook 엑셀 가져오기
        // 엑셀 validation확인 후 엑셀 확장자형식에따라 반환
        Workbook workbook = ExcelUtil.getWorkbook(file);

        //우선!! 원하는 시트의 데이터값들을 가져옴
        List<ArrayList> sheet1 = ExcelUtil.getWorksheet(workbook,1);
        model.addAttribute("dataList",sheet1);
        return "excelList";
    }
}
