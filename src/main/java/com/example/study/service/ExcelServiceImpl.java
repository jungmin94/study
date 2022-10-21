package com.example.study.service;

import com.example.study.common.ExcelUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Override
    public List<ArrayList> searchExcelList(MultipartFile file) throws Exception{

        //validation 후 엑셀 확장자형식에따라 반환
        Workbook workbook = ExcelUtil.getWorkbook(file);

        //맨 첫번째 시트의 데이터값들을 가져옴
        List<ArrayList> sheet1 = ExcelUtil.getWorksheet(workbook,0);

        return sheet1;
    }
}
