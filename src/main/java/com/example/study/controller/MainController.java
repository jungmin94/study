package com.example.study.controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

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
            List<ArrayList> dataList = new ArrayList<>();

            String fileName = file.getOriginalFilename();
            String extension = fileName.substring(fileName.indexOf('.')+1);

            Workbook workbook = null;
            //엑셀 확장자 형식에따라 Workbook instance 생성
            if(extension.equals("xlsx")){
                workbook = new XSSFWorkbook(file.getInputStream());
            }else if(extension.equals("xls")){
                workbook = new HSSFWorkbook(file.getInputStream());
            }

            // workbook의 첫번째 시트 가져오기
            Sheet worksheet = workbook.getSheetAt(0);

            // 모든 행(row)들을 조회한다
            for(Row row : worksheet) {
                //row의 모든 cell들을 순회한다
                Iterator<Cell> cellIterator = row.cellIterator();

                ArrayList<String> datas = new ArrayList();
                String value = "";

                //cell의 수만큼 조회한다
                while(cellIterator.hasNext()){
                    //cell의 타입하고 값을 가져온다
                    Cell cell = cellIterator.next();

                    //타입별로 값을 가져와 value에 저장한다.
                    switch (cell.getCellType()){
                        case FORMULA : //계산식처리
                                value =cell.getCellFormula();
                                break;
                        case NUMERIC : //숫자
                                value = (int)cell.getNumericCellValue()+"";
                                break;
                        case  STRING : //문자
                                value =cell.getStringCellValue();
                                break;
                        case  BLANK : //빈칸
                                value = cell.getBooleanCellValue()+"";
                                break;
                        case  ERROR : //에러
                                value = cell.getErrorCellValue()+"";
                                break;
                    }
                    //저장된 value값을 datas에 넣는다.
                    datas.add(value);
                }
                //한 row의 값들을 dataList에 넣는다
                dataList.add(datas);
            }

            model.addAttribute("dataList",dataList);

        return "excelList";
    }
}
