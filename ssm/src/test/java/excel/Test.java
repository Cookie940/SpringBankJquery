package excel;

import com.spring.pojo.City;
import com.spring.service.CityService;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.Region;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Test {
    public static void main(String[] args) throws IOException {
        ApplicationContext ctx = new ClassPathXmlApplicationContext("application.xml");
        CityService cityService =ctx.getBean("cityService",CityService.class);
        List<City> list= cityService.find();
        System.out.println(list.size());
        //List<Execls> list=new ArrayList<>();
        //1、创建HSSFWorkBook 对应一个excel
        HSSFWorkbook wb = new HSSFWorkbook();
        //1.5、生成excel中可能用到的单元格样式;
        //创建字体样式
        HSSFFont font = wb.createFont();
        font.setFontName("仿宋");//设置字体
        font.setFontHeightInPoints((short) 10);//设置字体大小
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//加粗


        //然后创建单元格样式style
        HSSFCellStyle style1 = wb.createCellStyle();
        style1.setFont(font);//将字体注入
        style1.setWrapText(true);// 自动换行
        style1.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 左右居中
        style1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 上下居中
        style1.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());// 设置单元格的背景颜色
        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style1.setBorderTop((short) 1);// 边框的大小
        style1.setBorderBottom((short) 1);
        style1.setBorderLeft((short) 1);
        style1.setBorderRight((short) 1);


        //2、生成一个sheet，对应excel的sheet，参数为excel中sheet显示的名字
        HSSFSheet sheet = wb.createSheet("城市列表");
        //3、设置sheet中每列的宽度，第一个参数为第几列，
        // 0为第一列；第二个参数为列的宽度，可以设置为0。
        // Test中有三个属性，因此这里设置三列，第0列设置宽度为0，第1~3列用以存放数据
        sheet.setColumnWidth(0, 0);
        sheet.setColumnWidth(1, 20*256);
        sheet.setColumnWidth(2, 20*256);
        sheet.setColumnWidth(3, 20*256);
        //4、生成sheet中一行，从0开始
        HSSFRow row = sheet.createRow(0);
        row.setHeight((short) 800);// 设定行的高度
        // 5、创建row中的单元格，从0开始
        HSSFCell cell;
        //我们第一列设置宽度为0，不会显示，因此第0个单元格不需要设置样式
        cell = row.createCell(1);//从第1个单元格开始，设置每个单元格样式
        cell.setCellValue("编号");//设置单元格中内容
        cell.setCellStyle(style1);//设置单元格样式

        cell = row.createCell(2);//第二个单元格
        cell.setCellValue("城市名");
        cell.setCellStyle(style1);

        cell = row.createCell(3);//第三个单元格
        cell.setCellValue("号码归属");
        cell.setCellStyle(style1);

        //6、输入数据
        int c=1;
        for(int i = 0; i <= list.size()-1; i++){
                HSSFRow rows = sheet.createRow(c);
                rows.setHeight((short) 800);
                cell = rows.createCell(2);
                City city = list.get(i);
                cell.setCellValue(city.getCname());
                cell.setCellStyle(style1);
                cell = rows.createCell(1);
                cell.setCellValue(city.getCid());
                cell.setCellStyle(style1);
                cell = rows.createCell(3);
                cell.setCellValue(city.getAreacode());
                cell.setCellStyle(style1);
                c++;
        }

        //7、如果需要单元格合并，有两种方式
        // 1、
        //sheet.addMergedRegion(new Region(1,(short)1,1,(short)11));//参数为(第一行，最后一行，第一列，最后一列)
        //2、
        //sheet.addMergedRegion(new CellRangeAddress(2, 3, 1, 1));//参数为(第一行，最后一行，第一列，最后一列)

        //8、输入excel
        FileOutputStream os = new FileOutputStream("D:/test.xls");
        wb.write(os);
        os.close();


    }

}
