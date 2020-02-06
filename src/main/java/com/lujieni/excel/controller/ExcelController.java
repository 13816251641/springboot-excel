package com.lujieni.excel.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import com.lujieni.excel.entity.StudentEntity;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;


/**
 * @Auther ljn
 * @Date 2020/2/5
 *
 */
@RestController
@RequestMapping("/excel")
@Slf4j
public class ExcelController {

    /**
     * 导出模板
     * @param response
     */
    @GetMapping("/export-template")
    public void exportTemplate(HttpServletResponse response){
        //导出操作
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("计算机一班学生","学生"), StudentEntity.class,new ArrayList<StudentEntity>());
        //告诉浏览器返回的是excel文件
        response.setHeader("content-type", "application/vnd.ms-excel");
        try {
            //设置excel的文件名称,这里一定要转码,因为服务端的编码是iso-8859-1
            response.setHeader("content-disposition", "attachment;filename="+new String("模板.xls".getBytes(StandardCharsets.UTF_8),StandardCharsets.ISO_8859_1));
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
            log.error("发生错误了",e);
        }
    }

    /**
     * 导入excel数据
     * excel中即使字段名顺序是乱的,只要字段名对也可以正常导入
     */
    @PostMapping("/import-excel")
    public void importExcel(MultipartFile file){
        ImportParams params = new ImportParams();
        params.setTitleRows(1);
        params.setHeadRows(1);
        /* 验证字段名是否符合,不符合会抛运行时异常 */
        params.setImportFields(new String[]{
               "学生姓名","学生性别","出生日期","进校日期"
        });
        try {
            List<StudentEntity> list = ExcelImportUtil.importExcel(file.getInputStream(), StudentEntity.class, params);
            System.out.println(list);
        } catch (Exception e) {
            e.printStackTrace();
            log.error("发生错误了",e);
        }
    }


}
