package com.example.xssf.utils;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;

import javax.swing.plaf.ColorUIResource;
import java.io.*;

import static org.apache.poi.ss.SpreadsheetVersion.EXCEL2007;

/**
 * @Author safin
 * @Date 2021/4/21 21:41
 * @Version 1.0
 */

public class PivotTableUtils {

    public static void main(String[] args) throws IOException {

        Workbook wb = new XSSFWorkbook(new FileInputStream(new File("C:\\Users\\szzft\\Desktop\\toushi-test2.xlsx")));

        try {
            toPivotTable(wb,2);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    /***
     * 保存为数据透视表
     * @param workbook
     * @param colCount
     * @throws Exception
     */
    private static void toPivotTable(Workbook workbook, int colCount) throws Exception {

        if (null == workbook ) {
            throw new Exception("数据集和元素实体不能为空!");
        }

        Sheet sheet1 = workbook.getSheetAt(0);//选定你要生成数据透视表的数据所在的sheet页

        // 创建数据透视sheet
        XSSFSheet pivotSheet = (XSSFSheet )workbook.createSheet();
        pivotSheet.setDefaultColumnWidth( 25);


        // 获取数据sheet的总行数
        int rowNum  = sheet1.getLastRowNum();
        System.out.println("==" + rowNum + "==");
        // 数据透视表生产的起点单元格位置
        CellReference ptStartCell = new CellReference("B1");
        AreaReference areaR=new AreaReference("B1:D"+ rowNum + 1,EXCEL2007);
        //从sheet1的选定数据范围内数据生成数据透视表
        XSSFPivotTable pivotTable = pivotSheet.createPivotTable(areaR, ptStartCell, sheet1);

        //透视表 列值
        pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2,"求和项:发送数量");
        //透视表 行标签
        pivotTable.addRowLabel(0);
        //透视表 行的值

        pivotTable.addRowLabel(2);

    /*    CTPivotFields ctPivotFields = pivotTable.getCTPivotTableDefinition().getPivotFields();
        CTPivotField fld ;

        for (int i = 0; i <= colCount; i++) {
            pivotTable.addRowLabel(i);//添加列
            fld = ctPivotFields.getPivotFieldList().get(i);
            fld.setOutline(false);//不以大纲形式显示
            fld.setCompact(false);//不以压缩形式显示

            for (int j = 0; j < rowNum -1; j++) {
                CTItems items = fld.getItems();
                items.getItemArray(j).unsetT();
                items.getItemArray(j).setX((long)j);
            }

            for (int k = rowNum -1; k > 1; k--) {
                //remove further items
                fld.getItems().removeItem(k);
            }
            fld.getItems().setCount(2);

            //build a cache definition which has shared elements for those items
            //<sharedItems><s v="Y"/><s v="N"/></sharedItems>
            pivotTable.getPivotCacheDefinition().getCTPivotCacheDefinition().getCacheFields().getCacheFieldList().get(i).getSharedItems().addNewS().setV("Y");
            pivotTable.getPivotCacheDefinition().getCTPivotCacheDefinition().getCacheFields().getCacheFieldList().get(i).getSharedItems().addNewS().setV("N");

            fld.setDefaultSubtotal(false);
        }

        pivotTable.getCTPivotTableDefinition().setMergeItem(false);//合并相同的单元格
        pivotTable.getCTPivotTableDefinition().setRowGrandTotals(true);//显示总计*/


        workbook.setSheetName(1,"pivotTable");
        workbook.write(new FileOutputStream("C:\\codes\\xssf\\" + System.currentTimeMillis() + "-pivotTable.xlsx"));
    }
}
