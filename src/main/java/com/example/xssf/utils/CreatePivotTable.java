package com.example.xssf.utils;

import com.spire.xls.*;
import com.spire.xls.core.IPivotTable;
import org.apache.poi.hslf.util.SystemTimeUtils;

/**
 * @Author safin
 * @Date 2021/4/18 16:08
 * @Version 1.0
 */

public class CreatePivotTable {
    public static void main(String[] args) {

        //加载示例文档
        Workbook workbook = new Workbook();
        workbook.loadFromFile("C:\\Users\\szzft\\Desktop\\toushi-test2.xlsx");

        //获取第一个工作表
        Worksheet sheet = workbook.getWorksheets().get(0);

        //为需要汇总和创建分析的数据创建缓存
        CellRange dataRange = sheet.getCellRange("B1:F76");
        PivotCache cache = workbook.getPivotCaches().add(dataRange);

        //使用缓存创建数据透视表，并指定透视表的名称以及在工作表中的位置
        PivotTable pt = sheet.getPivotTables().add("Pivot Table", sheet.getCellRange("H4"), cache);

        //添加行字段
        PivotField pf = null;
        if (pt.getPivotFields().get("一级部门名称") instanceof PivotField) {
            pf = (PivotField) pt.getPivotFields().get("一级部门名称");
        }
        pf.setAxis(AxisTypes.Row);
        /*PivotField pf2 =null;
        if (pt.getPivotFields().get("商品") instanceof PivotField){
            pf2= (PivotField) pt.getPivotFields().get("商品");
        }
        pf2.setAxis(AxisTypes.Row);
*/
        //添加值字段
        pt.getDataFields().add(pt.getPivotFields().get("发送数量"), "求和项：发送数量", SubtotalTypes.Sum);

        //设置透视表样式
        pt.setBuiltInStyle(PivotBuiltInStyles.PivotStyleMedium12);

        //生成透视图
        //获取该工作表中数据透视表
        IPivotTable pivotTable = sheet.getPivotTables().get(0);


        //保存文档
        workbook.saveToFile("C:\\codes\\xssf\\toushibiao.xlsx", ExcelVersion.Version2013);

        System.out.println("生成透视表成功");


        //加载包含透视表的Excel文档
        Workbook wb = new Workbook();
        wb.loadFromFile("C:\\codes\\xssf\\toushibiao.xlsx");

        //根据数据透视表创建数据透视图到第二个工作表
        Chart chart = workbook.getWorksheets().get(0).getCharts().add(ExcelChartType.ColumnClustered, pivotTable);
        //设置图表位置
        chart.setTopRow(2);
        chart.setLeftColumn(11);
        chart.setBottomRow(15);
        chart.setRightColumn(21);
        //设置图表标题
        chart.setChartTitle("汇总统计");

        //保存文档
        workbook.saveToFile("C:\\codes\\xssf\\"  + System.currentTimeMillis() + "-数据透视图.xlsx", ExcelVersion.Version2013);
        workbook.dispose();
        System.out.println("生成透视图成功");
    }
}
