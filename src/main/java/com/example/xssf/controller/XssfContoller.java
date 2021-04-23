package com.example.xssf.controller;

import com.example.xssf.utils.ExcelChartUtil22;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author safin
 * @Date 2021/3/23 22:10
 * @Version 1.0
 */
@RestController
@RequestMapping("xssf")
public class XssfContoller {
    @RequestMapping(value = "/getXssf", method = {RequestMethod.GET, RequestMethod.POST})
    @ResponseBody
    public String getXssf(@RequestBody String param) {
        System.out.println("展示图形------- param:" + param);
        return "哈哈";
    }

    @RequestMapping("/getExcel")
    @ResponseBody
    public String getExcel(MultipartFile file) throws IOException {
        InputStream inputStream = null;
        try {

            inputStream = file.getInputStream();//获取前端传递过来的文件对象，存储在“inputStream”中
            String fileName = file.getOriginalFilename();//获取文件名

            Workbook workbook = null; //用于存储解析后的Excel文件

            //判断文件扩展名为“.xls还是xlsx的Excel文件”,因为不同扩展名的Excel所用到的解析方法不同
            String fileType = fileName.substring(fileName.lastIndexOf("."));
            if (".xls".equals(fileType)) {
                workbook = new HSSFWorkbook(inputStream);//HSSFWorkbook专门解析.xls文件
            } else if (".xlsx".equals(fileType)) {
                workbook = new XSSFWorkbook(inputStream);//XSSFWorkbook专门解析.xlsx文件
            }

            ArrayList<ArrayList<Object>> list = new ArrayList<>();

            Sheet sheet; //工作表
            Row row;      //行
            Cell cell;    //单元格

            sheet = workbook.getSheet("sheet1");
            row = sheet.getRow(sheet.getFirstRowNum());

            ArrayList<String> fldNameArr = new ArrayList<>();
            fldNameArr.add("value1");
            fldNameArr.add("value2");
//            fldNameArr.add("value3");
//            fldNameArr.add("value4");
//            fldNameArr.add("value5");
//            fldNameArr.add("value6");
//            System.out.println("fldNameArr：" + fldNameArr.toString());

            List<String> titleArr = new ArrayList<>();
       /*     for (int k = row.getFirstCellNum() + 1; k < row.getLastCellNum(); k++) {
                cell = row.getCell(k);
                if (k == row.getFirstCellNum() + 1 || k == row.getLastCellNum() -1){
                    titleArr.add(cell.getStringCellValue());
                    fldNameArr.add("value" + k);
                }

            }*/
            titleArr.add(row.getCell(row.getFirstCellNum() + 1).getStringCellValue());
            titleArr.add(row.getCell(row.getLastCellNum() - 1).getStringCellValue());
            System.out.println("titleArr：" + titleArr.toString());

            List<Map<String, Object>> dataList = new ArrayList<Map<String, Object>>();
            //循环行  sheet.getPhysicalNumberOfRows()是获取表格的总行数
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                //System.out.println("第" + (i + 1) + "行内容:");
                row = sheet.getRow(i); // 取出第i行  getRow(index) 获取第(index+1)行

                Map<String, Object> dataMap = new HashMap<String, Object>();
                dataMap.put("value1", row.getCell(row.getFirstCellNum() + 1).getStringCellValue());
//                dataMap.put("value2", row.getCell(row.getFirstCellNum()).getNumericCellValue());
//                dataMap.put("value3", row.getCell(row.getFirstCellNum() + 3).getNumericCellValue());
//                dataMap.put("value4", row.getCell(row.getFirstCellNum() + 4).getNumericCellValue());
//                dataMap.put("value2", new DecimalFormat().format(row.getCell(row.getLastCellNum() - 1).getNumericCellValue() * 100) + "%" );

                dataMap.put("value2", row.getCell(row.getLastCellNum() - 1).getNumericCellValue());

                dataList.add(dataMap);
            }
            System.out.println("dataList：" + dataList);

          /*  //循环遍历，获取数据
            for (int j = sheet.getFirstRowNum() + 1; j <=sheet.getLastRowNum(); j++) {
                //从有数据的第二行开始遍历
                row = sheet.getRow(j);
                Map<String, Object> tempDataMap = new HashMap<String, Object>();
                if (row != null ) {
                    for (int k = row.getFirstCellNum() + 1; k <row.getLastCellNum(); k++) {
                        //这里需要注意的是getLastCellNum()的返回值为“下标+1”
                        cell = row.getCell(k);
                        if (k == row.getFirstCellNum() + 1 ){
                            tempDataMap.put("value" + k, cell.getStringCellValue());
                        }
                        else if (k == row.getLastCellNum()-1)
                        {
                            tempDataMap.put("value" + k, cell.getNumericCellValue());
                        }else {
                            continue;
                        }
                    }

                }
                dataList.add(tempDataMap);
            }*/

            System.out.println("dataList:" + dataList);
            System.out.println("fldNameArr：" + fldNameArr.toString());

            ExcelChartUtil22 ecu = new ExcelChartUtil22();

            ecu.setWb(new SXSSFWorkbook());
            // 创建柱状图
            ecu.createBarChart(titleArr, fldNameArr, dataList);


            FileOutputStream out = new FileOutputStream(new File("c:\\codes\\xssf\\" + System.currentTimeMillis() + ".xlsx"));
            ecu.getWb().write(out);
            out.close();


            //将内容保存为输入流
//            ByteArrayInputStream in = null;
//            try {
//                ByteArrayOutputStream os = new ByteArrayOutputStream();
//                ecu.getWb().write(os);
//
//                byte[] b = os.toByteArray();
//                in = new ByteArrayInputStream(b);
//                //上传到FastDFS服务器
//                //ExcelChartUtil.uploadFastDFS(in,"01");
//
//                os.close();
//
//            } catch (IOException e) {
//                 e.printStackTrace();
//                //logger.error("ExcelUtils getExcelFile error:{}",e.toString());
//
//                return null;
//            }
            //return in;


        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != inputStream) {
                inputStream.close();
            }
        }
        return "success";
    }

}
