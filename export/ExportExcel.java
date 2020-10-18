package com.javase.export;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.*;

public class ExportExcel {

    public static void main(String[] args) {

        ExportExcel exportExcel = new ExportExcel();

        try {
            List<Student> studentList = exportExcel.downloadExcel();
            System.out.println(studentList);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private List<Student> downloadExcel() throws Exception{

        List<Student> studentList = new ArrayList<>();
        Student student;

        for (int i = 0; i < 10; i++) {
            student = new Student();
            student.setId(i + 1);
            student.setStudentNo("123456");
            student.setStudentName("小雪");
            student.setSex("女");
            student.setAddress("北京市大兴区");
            student.setCreateTime(new Date());
            student.setUpdateTime(new Date());
            studentList.add(student);
        }

        System.out.println("studentList:" + studentList);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(); //创建一个Excel
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("学生信息管理");//创建一个sheet
        xssfSheet.setDefaultColumnWidth(2);//默认列宽
        xssfSheet.setDefaultRowHeightInPoints((float) 11.25);//默认行高
        Header header = xssfSheet.getHeader();//设置sheet的头


        XSSFCellStyle titleStyle = xssfWorkbook.createCellStyle();//设置标题样式
        XSSFFont xssfTitleFont = xssfWorkbook.createFont();//设置字体
        xssfTitleFont.setFontHeightInPoints((short)16); //设置单元字体高度
        xssfTitleFont.setFontName("楷体");
        titleStyle.setFont(xssfTitleFont);
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);


        XSSFCellStyle headStyle = xssfWorkbook.createCellStyle();//设置表头的样式
//        headStyle.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont xssfHeadFont = xssfWorkbook.createFont();//设置字体
        xssfHeadFont.setFontHeightInPoints((short) 9);//设置单元字体高度
        xssfHeadFont.setFontName("仿宋");
        xssfHeadFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        headStyle.setFont(xssfHeadFont);
        headStyle.setAlignment(HorizontalAlignment.LEFT);
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headStyle.setBorderBottom(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//设置填充方案
        headStyle.setFillForegroundColor(new XSSFColor(new Color(192,192,192))); //设置自定义填充颜色


        XSSFCellStyle dataStyle = xssfWorkbook.createCellStyle(); //设置数据样式
//        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont xssfDataFont = xssfWorkbook.createFont();
        xssfDataFont.setFontHeightInPoints((short)9); //设置单元字体高度
        xssfDataFont.setFontName("宋体");
        dataStyle.setFont(xssfDataFont);
        dataStyle.setAlignment(HorizontalAlignment.LEFT);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        dataStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        dataStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        dataStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);


        String []tableHeader = {"No.","学号","姓名","性别","家庭住址","创建时间","更新时间"};
        List<String> headDataList = new ArrayList<>(Arrays.asList(tableHeader));

//        short headCellNumber=(short)tableHeader.length;//表的列数

        try {
            //根据是否取出数据，设置header信息
            if(studentList.size() == 0){
                header.setCenter("查无资料");
                System.out.println("null....");
            }else{
                System.out.println("ok....");
//                header.setCenter("学生信息管理");

                setTitleCombineCell(xssfSheet);//设置标题合并单元格
                setTitleValue(xssfSheet, titleStyle);//设置标题信息

                int rowNum = 12;//表头行号
                int cellNum = 1;//列号

                XSSFRow row = xssfSheet.createRow(rowNum);//创建第rowNum行
//                XSSFCell cell = row.createCell(cellNum);//创建第rowNum行第cellNum列

                setTableCombineCell(xssfWorkbook, xssfSheet, rowNum);//设置表头合并单元格
                //获取所有的合并单元格
                List<CellRangeAddress> combineCellList = getCombineCellList(xssfSheet);
                setHeadValue(row, combineCellList, headDataList, cellNum);
                setRegionStyle(row, headStyle);//为表头设置样式

                // 给Excel填充数据
                for(int i = 0; i < studentList.size(); i++){
                    row = xssfSheet.createRow((short) (i + 13));//创建第i+13行
                    rowNum = row.getRowNum();
                    cellNum = 1;
                    //为当前行设置合并单元格
                    setTableCombineCell(xssfWorkbook, xssfSheet, rowNum);
                    //获取所有的合并单元格
                    combineCellList = getCombineCellList(xssfSheet);
                    setDataValue(row, combineCellList, studentList.get(i), cellNum);
                    //为当前行设置样式
//                    XSSFCellStyle dataStyle = xssfWorkbook.createCellStyle();
                    setRegionStyle(row, dataStyle);

//                    if (i==0) xssfSheet.autoSizeColumn(i);

                    // POI分页符有BUG，必须在模板文件中插入一个分页符，然后再此处删除预设的分页符；最后在下面重新设置分页符。

//                    xssfSheet.setAutobreaks(false);
//                    int iRowBreaks[] = xssfSheet.getRowBreaks();
//                    xssfSheet.removeRowBreak(3);
//                    xssfSheet.removeRowBreak(4);
//                    xssfSheet.removeRowBreak(5);
//                    xssfSheet.removeRowBreak(6);



                    if (i%10 == 0){
                        xssfSheet.setRowBreak(i+22); //在第startRow行设置分页符
                    }


                }

//                row = xssfSheet.createRow((short) (row.getRowNum() + 2));//创建第i+13行
//                row.getRowStyle().setHidden(true);

                CTCol col = xssfSheet.getCTWorksheet().getColsArray(0).addNewCol();
                col.setMin(25);
                col.setMax(16384); // the last column (1-indexed)
                col.setHidden(true);
//                CTRow ctRow = xssfSheet.getCTWorksheet().getSheetData().getRowArray(0).;
//                ctRow.set
//                CTWorksheet ctWorksheet = xssfSheet.getCTWorksheet();
//                xssfSheet.getCTWorksheet().get
                Iterator<Row> rowIterator = xssfSheet.iterator();
                while(rowIterator.hasNext()) {
                    Row tempRow = rowIterator.next();
                    System.out.println("tempRow.getRowNum():" + tempRow.getRowNum());
                    System.out.println("studentList.size() + 15:" + studentList.size() + 15);
                    if (tempRow.getRowNum() > studentList.size() + 15){
                        if (tempRow.getZeroHeight()) {
                            tempRow.setZeroHeight(false);
                        }
                    }

                }

                    // 固定第一行标题
//                xssfSheet.createFreezePane(1, 1, 1, 1);
//                xssfSheet.createSplitPane(2, 3, 5, 7, 9);
//
//                //分页
//                xssfSheet.setFitToPage(true);
//
//                xssfSheet.setRowBreak(25);
//                xssfSheet.setColumnBreak(25);



//                xssfSheet.setAutobreaks(true);
//                xssfSheet.setFitToPage(true);
//                PrintSetup printSetup = xssfSheet.getPrintSetup();
//                printSetup.setFitHeight((short)0);
//                printSetup.setFitWidth((short)1);

                // 调整宽度
//                xssfSheet.setColumnWidth(24, 10 * 100);



            }
        } catch (Exception e) {
            e.printStackTrace();
        }
//        outputSetting("注文一覧照会.xlsx", response, xssfWorkbook);C:\\Users\\user-pc\\Desktop
        //C:\\Users\\user-pc\\Desktop\\注文一覧照会6.xlsx
        System.out.println("写入数据完成，正在导出。。。");

        OutputStream op = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\学生信息管理一览.xlsx");
        xssfWorkbook.write(op);
        op.close();
//        String fileName = "注文一覧照会8.xlsx";
//        outputSetting(fileName, response, xssfWorkbook);

        System.out.println("导出文件成功，响应数据到前台。。。");
        return studentList;
    }

    private void setTitleValue(XSSFSheet xssfSheet, XSSFCellStyle titleStyle) {
        SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd");

        XSSFRow row = xssfSheet.createRow(1);
        XSSFCell cell = row.createCell(1);
        cell.setCellValue("学生信息管理系统");

        row = xssfSheet.createRow(2);
        row.setHeight((short)(24*20));
        cell = row.createCell(10);
        cell.setCellValue("学生信息一览");
        cell.setCellStyle(titleStyle);

        row = xssfSheet.createRow(3);
        cell = row.createCell(16);
        cell.setCellValue("查询日期:");
        cell = row.createCell(19);
        cell.setCellValue(format.format(new Date()));

        row = xssfSheet.createRow(4);
        cell = row.createCell(16);
        cell.setCellValue("操作人员:");
        cell = row.createCell(19);
        cell.setCellValue("Admin");
        cell = row.createCell(21);
        cell.setCellValue("管理员");

        row = xssfSheet.createRow(6);
        cell = row.createCell(2);
        cell.setCellValue("姓名:");
        cell = row.createCell(4);
        cell.setCellValue("1234567");

        row = xssfSheet.createRow(7);
        cell = row.createCell(2);
        cell.setCellValue("性别:");
        cell = row.createCell(4);
        cell.setCellValue("男");
    }

    private void setTitleCombineCell(XSSFSheet xssfSheet) {
        /*
          参数1： 起始单元格的行数
          参数2：结束单元格的行数
          参数3： 起始单元格的列
          参数4： 结束单元格的列
         */
        int row = 1;
//        int col = 0;
        CellRangeAddress range=new CellRangeAddress(row, row, 1, 8);
        xssfSheet.addMergedRegion(range);

        row = 2;
        range=new CellRangeAddress(row, row, 10, 15);
        xssfSheet.addMergedRegion(range);

        row = 3;
        range=new CellRangeAddress(row, row, 16, 18);
        xssfSheet.addMergedRegion(range);
        range=new CellRangeAddress(row, row, 19, 23);
        xssfSheet.addMergedRegion(range);

        row = 4;
        range=new CellRangeAddress(row, row, 16, 18);
        xssfSheet.addMergedRegion(range);
        range=new CellRangeAddress(row, row, 19, 20);
        xssfSheet.addMergedRegion(range);
        range=new CellRangeAddress(row, row, 21, 23);
        xssfSheet.addMergedRegion(range);

//        row = 6;
        for (row = 6; row < 8; row++) {
            range=new CellRangeAddress(row, row, 2, 3);
            xssfSheet.addMergedRegion(range);
            range=new CellRangeAddress(row, row, 4, 6);
            xssfSheet.addMergedRegion(range);
        }

    }

    private void setHeadValue(XSSFRow row, List<CellRangeAddress> combineCellList, List<String> headDataList, int cellNum) {
        int tempNum = cellNum;
        for (String headData :
                headDataList) {
            tempNum = setTableValue(row, combineCellList, headData, tempNum);
        }
    }

    private void setDataValue(XSSFRow row, List<CellRangeAddress> combineCellList, Student student, int cellNum) {
        SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd");
        cellNum = setTableValue(row, combineCellList, String.valueOf(student.getId()), cellNum);
        cellNum = setTableValue(row, combineCellList, student.getStudentNo(), cellNum);
        cellNum = setTableValue(row, combineCellList, student.getStudentName(), cellNum);
        cellNum = setTableValue(row, combineCellList, student.getSex(), cellNum);
        cellNum = setTableValue(row, combineCellList, student.getAddress(), cellNum);
        cellNum = setTableValue(row, combineCellList, format.format(student.getCreateTime()), cellNum);
        setTableValue(row, combineCellList, format.format(student.getUpdateTime()), cellNum);
    }

    private int setTableValue(XSSFRow row, List<CellRangeAddress> combineCellList, String str, int cellNum) {

        XSSFCell cell = row.createCell((short) cellNum);//创建第i+13行第0列
        cell.setCellValue(str);//设置第i+1行第0列的值
        //判断是否为合并单元格
        Map<String, Object> combineCell = isCombineCell(cell, combineCellList);
        boolean flag = (boolean) combineCell.get("flag");
        if (flag) {
            int mergedRow = (int) combineCell.get("mergedRow");
            int mergedCol = (int) combineCell.get("mergedCol");
            System.out.printf("行:%s, 列:%s, 行数:%s, 列数:%s\n", row.getRowNum(), cellNum, mergedRow, mergedCol);
            cellNum = cellNum + mergedCol;
        }else {
            cellNum++;
        }

        return cellNum;
    }

    private void setTableCombineCell(XSSFWorkbook xssfWorkbook, XSSFSheet xssfSheet, int rowNum) {
        CellRangeAddress range = new CellRangeAddress(rowNum, rowNum, 2, 5);//学号
        xssfSheet.addMergedRegion(range);
        setMergedRegionStyle(xssfWorkbook, xssfSheet, range);

        range=new CellRangeAddress(rowNum, rowNum, 6, 8);//姓名
        xssfSheet.addMergedRegion(range);
        setMergedRegionStyle(xssfWorkbook, xssfSheet, range);

        range=new CellRangeAddress(rowNum, rowNum, 9, 11);//性别
        xssfSheet.addMergedRegion(range);
        setMergedRegionStyle(xssfWorkbook, xssfSheet, range);

        range=new CellRangeAddress(rowNum, rowNum, 12, 16);//家庭住址
        xssfSheet.addMergedRegion(range);
        setMergedRegionStyle(xssfWorkbook, xssfSheet, range);

        range=new CellRangeAddress(rowNum, rowNum, 17, 19);//创建时间
        xssfSheet.addMergedRegion(range);
        setMergedRegionStyle(xssfWorkbook, xssfSheet, range);

        range=new CellRangeAddress(rowNum, rowNum, 20, 22);//更新时间
        xssfSheet.addMergedRegion(range);
        setMergedRegionStyle(xssfWorkbook, xssfSheet, range);
    }

    private static void setMergedRegionStyle(XSSFWorkbook xssfWorkbook, XSSFSheet xssfSheet, CellRangeAddress range){
        RegionUtil.setBorderBottom(HSSFCellStyle.BORDER_THIN, range, xssfSheet, xssfWorkbook);
        RegionUtil.setBorderTop(HSSFCellStyle.BORDER_THIN, range, xssfSheet, xssfWorkbook);
        RegionUtil.setBorderLeft(HSSFCellStyle.BORDER_THIN, range, xssfSheet, xssfWorkbook);
        RegionUtil.setBorderRight(HSSFCellStyle.BORDER_THIN, range, xssfSheet, xssfWorkbook);
    }

    private static void setRegionStyle(XSSFRow row, XSSFCellStyle cellStyle) {
        System.out.println("现在在第" + row.getRowNum() + "行");
        //为每个单元格设置边框，问题就解决了
        XSSFCell cell;
        for(int i = 1; i<= 22; i++){

            cell = row.getCell(i);


            if (cell != null) {
                System.out.println("现在在第" + row.getRowNum() + "行, 第" + cell.getColumnIndex() + "列");

                System.out.println("cell.getRowIndex() :" + cell.getRowIndex());
                System.out.println("cell.getColumnIndex():" + cell.getColumnIndex());
                if (row.getRowNum() > 12 && (cell.getColumnIndex() == 1 | cell.getColumnIndex() > 16)){
                    cellStyle.setAlignment(HorizontalAlignment.RIGHT);//HorizontalAlignment.RIGHT
                }else{
                    cellStyle.setAlignment(HorizontalAlignment.LEFT);//CellStyle.ALIGN_LEFT  XSSFCellStyle.ALIGN_CENTER
                }
                cell.setCellStyle(cellStyle);
            }
        }
    }

    //获取合并单元格集合
    private static List<CellRangeAddress> getCombineCellList(XSSFSheet sheet) {
        List<CellRangeAddress> list = new ArrayList<>();
        //获得一个 sheet 中合并单元格的数量
        int sheetmergerCount = sheet.getNumMergedRegions();
        //遍历所有的合并单元格
        for(int i = 0; i<sheetmergerCount;i++)
        {
            //获得合并单元格保存进list中
            CellRangeAddress ca = sheet.getMergedRegion(i);
            list.add(ca);
        }
        return list;
    }

    /**
     * 判断cell是否为合并单元格，是的话返回合并行数和列数（只要在合并区域中的cell就会返回合同行列数，但只有左上角第一个有数据）
     * @param listCombineCell  上面获取的合并区域列表
     * @param cell 当前单元格
     * @return 单元格集合map对象
     */
    private static Map<String,Object> isCombineCell(XSSFCell cell, List<CellRangeAddress> listCombineCell) {
        int firstC, lastC, firstR, lastR, mergedRow, mergedCol;
//        String cellValue = null;
        Boolean flag = false;
        Map<String,Object> result=new HashMap<>();
        result.put("flag", flag);
        for(CellRangeAddress ca:listCombineCell)
        {
            //获得合并单元格的起始行, 结束行, 起始列, 结束列
            firstC = ca.getFirstColumn();
            lastC = ca.getLastColumn();
            firstR = ca.getFirstRow();
            lastR = ca.getLastRow();
            //判断cell是否在合并区域之内，在的话返回true和合并行列数
            if(cell.getRowIndex() >= firstR && cell.getRowIndex() <= lastR)
            {
                if(cell.getColumnIndex() >= firstC && cell.getColumnIndex() <= lastC)
                {
                    flag = true;
                    mergedRow=lastR-firstR+1;
                    mergedCol=lastC-firstC+1;
                    result.put("flag", flag);
                    result.put("mergedRow",mergedRow);
                    result.put("mergedCol",mergedCol);
                    break;
                }
            }
        }
        return result;
    }


    /*
    //固定配置
    private void outputSetting(String fileName, HttpServletResponse response, XSSFWorkbook xssfWorkbook) {

        //private void download(String path, HttpServletResponse response) {
    	/*
    	try {
			// path是指欲下载的文件的路径。
			File file = new File(path);
			// 取得文件名。
			String filename = file.getName();
			// 以流的形式下载文件。
			InputStream fis = new BufferedInputStream(new FileInputStream(path));
			byte[] buffer = new byte[fis.available()];
			fis.read(buffer);
			fis.close();
			// 清空response
			response.reset();
			// 设置response的Header
			response.addHeader("Content-Disposition", "attachment;filename=" + new String(filename.getBytes()));
			response.addHeader("Content-Length", "" + file.length());


			OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
			response.setContentType("application/vnd.ms-excel;charset=gb2312");

			toClient.write(buffer);
			toClient.flush();
			toClient.close();
		} catch (IOException ex) {
			ex.printStackTrace();
		}
    	*/

//        ServletActionContext servletActionContext =
//        applicationContext.con

//        HttpServletResponse response = null;//创建一个HttpServletResponse对象
       /*
        OutputStream out = null;//创建一个输出流对象

        try {
//            response = ServletActionContext.getResponse();//初始化HttpServletResponse对象

            out = response.getOutputStream();// 得到输出流
            response.setHeader("Content-disposition","attachment; filename="+new String(fileName.getBytes(),"UTF-8"));//filename是下载的xls的名ISO-8859-1
            response.setContentType("application/msexcel;charset=UTF-8");//设置类型
            response.setHeader("Pragma","No-cache");//设置头
            response.setHeader("Cache-Control","no-cache");//设置头
            response.setDateHeader("Expires", 0);//设置日期头


            xssfWorkbook.write(out);
            out.flush();
            out.close();

//            xssfWorkbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            try{
                if(out!=null){
                    out.close();
                }
            }catch(IOException e){
                e.printStackTrace();
            }
        }
    }

    */
}
