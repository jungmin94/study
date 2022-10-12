package com.example.study.controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import java.io.IOException;
import java.util.*;

@Controller
public class MainController {

    @RequestMapping("/")
    public String index() {
        return "index";
    }

    @RequestMapping("/excel/selectExcel")
    public String selectExcel(){
        return "excel";
    }

    @RequestMapping("/excel/readExcel")
    public String readExcel(@RequestParam("file") MultipartFile file, Model model)
            throws IOException{
        List<Map<String,Object>> dataList = new ArrayList<>();

        String fileName = file.getOriginalFilename();
        String extension = fileName.substring(fileName.indexOf('.')+1);

        Workbook workbook = null;
        //엑셀 확장자에따라 Workbook instance 생성
        if(extension.equals("xlsx")){
            workbook = new XSSFWorkbook(file.getInputStream());
        }else if(extension.equals("xls")){
            workbook = new HSSFWorkbook(file.getInputStream());
        }

        System.out.println("workbook : "+workbook);

        // workbook의 첫번째 시트 가져오기
        Sheet worksheet = workbook.getSheetAt(0);

        // 모든 행(row)들을 조회한다
        for(Row row : worksheet) {
            //row의 모든 cell들을 순회한다
            Iterator<Cell> cellIterator = row.cellIterator();

            //cell의 수만큼 조회한다
            while(cellIterator.hasNext()){
                //cell의 타입하고 값을 가져온다
                Cell cell = cellIterator.next();

//                if(cell.getCellType().toString()=="NUMERIC"){
//                    datas.put(cellIterator.le,cell.getNumericCellValue());
//                }else if(cell.getCellType().toString()=="STRING"){
//                    dataList.add(cell.getStringCellValue());
//                }else{
//                    dataList.add(cell.getDateCellValue());
//                }

            }
            System.out.println("dataList : " + dataList);
        }
            model.addAttribute("dataList",dataList);
        return "excelList";
    }
}
