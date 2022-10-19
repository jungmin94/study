package com.example.study.controller;

import com.example.study.common.ExcelUtill;
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

    @RequestMapping("/readExcel")
    public String readExcel(@RequestParam("file") MultipartFile file, Model model)
            throws IOException {
        List<ArrayList> dataList = new ArrayList<>();

        //getWorkbook 엑셀가져오기
        Workbook workbook = ExcelUtill.getWorkbook(file);


        //getSheet (파라미터 인덱스)
        // workbook의 첫번째 시트 가져오기
        Sheet worksheet = workbook.getSheetAt(0);

        // 모든 행(row)들을 조회한다
        for(Row row : worksheet) {
            //row의 모든 cell들을 순회한다
            Iterator<Cell> cellIterator = row.cellIterator();

            ArrayList<String> datas = new ArrayList<>();

            //cell의 수만큼 조회한다
            while(cellIterator.hasNext()){
                //cell의 타입하고 값을 가져온다
                Cell cell = cellIterator.next();

                String value = ExcelUtill.getCellValue(cell);

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
