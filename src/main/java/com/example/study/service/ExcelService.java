package com.example.study.service;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.List;

public interface ExcelService {

    public List<ArrayList> searchExcelList(MultipartFile file) throws Exception;
}
