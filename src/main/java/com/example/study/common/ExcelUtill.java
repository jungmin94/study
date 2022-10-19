package com.example.study.common;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelUtill {

    public static Workbook getWorkbook(MultipartFile file) throws IOException {
        String fileName = file.getOriginalFilename();
        String extension = fileName.substring(fileName.indexOf('.')+1);

        if(!extension.equals("xlsx") && !extension.equals("xls")){
            throw new IOException("엑셀파일만 업로드 해주세요");
        }

        Workbook workbook = null;
        //엑셀 확장자 형식에따라 Workbook instance 생성
        if(extension.equals("xlsx")){
            workbook = new XSSFWorkbook(file.getInputStream());
        }else if(extension.equals("xls")){
            workbook = new HSSFWorkbook(file.getInputStream());
        }
        return workbook;
    }

    public static String getCellValue(Cell cell) {
        String value="";
            //타입별로 값을 가져와 value에 저장한다.
            switch (cell.getCellType()) {
                case FORMULA: //계산식처리
                    value = cell.getCellFormula();
                    break;
                case NUMERIC: //숫자
                    value = (int) cell.getNumericCellValue() + "";
                    break;
                case STRING: //문자
                    value = cell.getStringCellValue();
                    break;
                case BLANK: //빈칸
                    value = cell.getBooleanCellValue() + "";
                    break;
                case ERROR: //에러
                    value = cell.getErrorCellValue() + "";
                    break;
            }

        return value;
    }
}
