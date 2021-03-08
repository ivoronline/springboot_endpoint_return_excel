package com.ivoronline.springboot_endpoint_return_excel.controllers;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

@Controller
public class MyController {

  @ResponseBody
  @GetMapping("/GetExcel")
  public ResponseEntity<StreamingResponseBody> getExcel() {

    //CREATE EXCEL FILE
    Workbook workBook = new XSSFWorkbook();

    //CREATE TABLE
    Sheet    sheet = workBook.createSheet("My Sheet");
             sheet.setColumnWidth(0, 10 * 256);        //10 characters wide

    //CREATE ROW
    Row      row = sheet.createRow(0);

    //CREATE CELL
    Cell     cell = row.createCell(0);
             cell.setCellValue("Hello World");

    //DOWNLOAD EXCEL
    return ResponseEntity
      .ok()
      .contentType(MediaType.APPLICATION_OCTET_STREAM)
      .header(HttpHeaders.CONTENT_DISPOSITION, "inline;filename=\"test.xlsx\"")
      .body(workBook::write);
  }

}
