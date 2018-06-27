package com.guozhaotong;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelUtil {
    private boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    private boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    /**
     * 总行数
     */
    private int totalRows = 0;

    /**
     * 总列数
     */
    private int totalCells = 0;

    /**
     * 错误信息
     */
    private String errorInfo;

    /**
     * 构造方法
     */
    private ExcelUtil() {

    }

    /**
     * 得到总行数
     */
    public int getTotalRows() {
        return totalRows;
    }

    /**
     * 得到总列数
     */
    private int getTotalCells() {
        return totalCells;
    }

    /**
     * 得到错误信息
     */
    public String getErrorInfo() {
        return errorInfo;
    }

    /**
     * 验证excel文件
     */
    private boolean validateExcel(String filePath) {
        // 检查文件名是否为空或者是否是Excel格式的文件/
        ExcelUtil excelUtil = new ExcelUtil();
        if (filePath == null || !(excelUtil.isExcel2003(filePath) || excelUtil.isExcel2007(filePath))) {
            errorInfo = "文件名不是excel格式";
            return false;
        }

        // 检查文件是否存在/
        File file = new File(filePath);
        if (file == null || !file.exists()) {
            errorInfo = "文件不存在";
            return false;
        }
        return true;
    }

    /**
     * 根据文件名读取excel文件
     */
    private List<List<Object>> readOneSheet(String filePath, int sheetIndex) {
        ExcelUtil excelUtil = new ExcelUtil();
        List<List<Object>> dataLst = new ArrayList<>();
        InputStream is = null;
        try {
            //验证文件是否合法
            if (!validateExcel(filePath)) {
                System.out.println(errorInfo);
                return null;
            }

            // 判断文件的类型，是2003还是2007
            boolean isExcel2003 = true;
            if (excelUtil.isExcel2007(filePath)) {
                isExcel2003 = false;
            }

            // 调用本类提供的根据流读取的方法
            File file = new File(filePath);
            is = new FileInputStream(file);
            dataLst = readOneSheet(is, isExcel2003, sheetIndex);
            is.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        // 返回最后读取的结果
        return dataLst;
    }

    private List<List<List<Object>>> readAllSheets(String filePath) {
        ExcelUtil excelUtil = new ExcelUtil();
        List<List<List<Object>>> dataLst = new ArrayList<>();
        InputStream is = null;
        try {
            //验证文件是否合法
            if (!validateExcel(filePath)) {
                System.out.println(errorInfo);
                return null;
            }

            // 判断文件的类型，是2003还是2007
            boolean isExcel2003 = true;
            if (excelUtil.isExcel2007(filePath)) {
                isExcel2003 = false;
            }

            // 调用本类提供的根据流读取的方法
            File file = new File(filePath);
            is = new FileInputStream(file);
            dataLst = readAllSheets(is, isExcel2003);
            is.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        // 返回最后读取的结果
        return dataLst;
    }

    /**
     * 根据流读取Excel文件
     */
    private List<List<Object>> readOneSheet(InputStream inputStream, boolean isExcel2003, int sheetIndex) {
        List<List<Object>> dataLst = null;
        try {
            Workbook wb;
            if (isExcel2003) {
                wb = new HSSFWorkbook(inputStream);
            } else {
                wb = new XSSFWorkbook(inputStream);
            }

            dataLst = read(wb, sheetIndex);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return dataLst;
    }

    private List<List<List<Object>>> readAllSheets(InputStream inputStream, boolean isExcel2003) {
        List<List<List<Object>>> dataLst = new ArrayList<>();
        try {
            Workbook wb;
            if (isExcel2003) {
                wb = new HSSFWorkbook(inputStream);
            } else {
                wb = new XSSFWorkbook(inputStream);
            }
            int sheetNumber = wb.getNumberOfSheets();
            for(int i = 0; i < sheetNumber; i++){
                List<List<Object>> oneSheet = read(wb, i);
                dataLst.add(oneSheet);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return dataLst;
    }

    /**
     * 读取数据
     */
    private List<List<Object>> read(Workbook wb, int sheetIndex) {
        List<List<Object>> dataLst = new ArrayList<>();

        // 得到一个shell
        Sheet sheet = wb.getSheetAt(sheetIndex);

        // 得到Excel的行数
        this.totalRows = sheet.getPhysicalNumberOfRows();

        // 得到Excel的列数
        if (this.totalRows >= 1 && sheet.getRow(0) != null) {
            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }

        // 循环Excel的行
        for (int r = 0; r < this.totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            List<Object> rowLst = new ArrayList<>();

            // 循环Excel的列
            for (int c = 0; c < this.getTotalCells(); c++) {
                Cell cell = row.getCell(c);
                Object cellValue = null;
                if (null != cell) {
                    // 以下是判断数据的类型
                    switch (cell.getCellType()) {
                        case 0: // 数字
                            cellValue = cell.getNumericCellValue();
                            break;

                        case 1: // 字符串
                            cellValue = cell.getStringCellValue();
                            break;
                        case 2: // 公式
                            cellValue = cell.getCellFormula();
                            break;

                        case 3: // 空值
                            cellValue = "";
                            break;

                        case 4: // Boolean
                            cellValue = cell.getBooleanCellValue();
                            break;

                        case 5: // 故障
                            cellValue = "非法字符";
                            break;

                        default:
                            cellValue = "未知类型";
                            break;
                    }
                }
                rowLst.add(cellValue);
            }

            // 保存第r行的第c列
            dataLst.add(rowLst);
        }
        return dataLst;
    }

    private void insertData2Sheet(HSSFSheet sheet, List<List<String>> dataList) {
        for (int i = 0; i < dataList.size(); i++) {
            HSSFRow row = sheet.createRow(i);
            for (int j = 0; j < dataList.get(i).size(); j++) {
                row.createCell(j).setCellValue(dataList.get(i).get(j));
            }
        }
    }

    private void insertData2Sheet(XSSFSheet sheet, List<List<String>> dataList) {
        for (int i = 0; i < dataList.size(); i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < dataList.get(i).size(); j++) {
                row.createCell(j).setCellValue(dataList.get(i).get(j));
            }
        }
    }

    private void writeWorkbook2File(String filePath, Workbook workBook) {
        File file = new File(filePath);
        //文件输出流
        FileOutputStream outStream = null;
        try {
            outStream = new FileOutputStream(file);
            workBook.write(outStream);
            outStream.flush();
            outStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void write2003(String fileName, String sheetName, List<List<String>> dataList) {

        // 声明一个工作薄
        HSSFWorkbook workBook = new HSSFWorkbook();
        // 生成一个表格
        HSSFSheet sheet = workBook.createSheet();
        workBook.setSheetName(0, sheetName);
        insertData2Sheet(sheet, dataList);
        writeWorkbook2File(fileName, workBook);
    }

    private void write2003(String fileName, List<List<String>> dataList) {
        write2003(fileName, "sheet0", dataList);
    }

    private void write2007(String fileName, String sheetName, List<List<String>> dataList) {

        // 声明一个工作薄
        XSSFWorkbook workBook = new XSSFWorkbook();
        // 生成一个表格
        XSSFSheet sheet = workBook.createSheet();
        workBook.setSheetName(0, sheetName);
        insertData2Sheet(sheet, dataList);
        writeWorkbook2File(fileName, workBook);
    }

    private void write2007(String fileName, List<List<String>> dataList) {
        write2007(fileName, "sheet0", dataList);
    }

    private void write(String fileName, String sheetName, List<List<String>> dataList) {
        if (isExcel2003(fileName)) {
            write2003(fileName, sheetName, dataList);
            return;
        }
        write2007(fileName, sheetName, dataList);
    }

    private void write(String fileName, List<List<String>> dataList) {
        write(fileName, "sheet0", dataList);
    }

    /**
     * 写多个工作簿的Excel
     */
    private void writeMultiSheet2003(String fileName, List<String> sheetName, List<List<List<String>>> dataList) {
        if (sheetName.size() != dataList.size()) {
            System.err.println("两个集合长度不一致，请检查！");
        }
        try {
            // 声明一个工作薄
            HSSFWorkbook workBook = new HSSFWorkbook();
            // 生成表格
            for (int j = 0; j < dataList.size(); j++) {
                HSSFSheet sheet = workBook.createSheet();
                workBook.setSheetName(j, sheetName.get(j));
                //插入需导出的数据
                for (int i = 0; i < dataList.get(j).size(); i++) {
                    HSSFRow row = sheet.createRow(i);
                    List<String> oneRowData = dataList.get(j).get(i);
                    for (int k = 0; k < oneRowData.size(); k++) {
                        row.createCell(k).setCellValue(oneRowData.get(k));
                    }
                }
            }
            writeWorkbook2File(fileName, workBook);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void writeMultiSheet2007(String fileName, List<String> sheetName, List<List<List<String>>> dataList) {
        if (sheetName.size() != dataList.size()) {
            System.err.println("两个集合长度不一致，请检查！");
        }
        try {
            XSSFWorkbook workBook = new XSSFWorkbook();
            for (int j = 0; j < dataList.size(); j++) {
                XSSFSheet sheet = workBook.createSheet();
                workBook.setSheetName(j, sheetName.get(j));
                for (int i = 0; i < dataList.get(j).size(); i++) {
                    XSSFRow row = sheet.createRow(i);
                    List<String> oneRowData = dataList.get(j).get(i);
                    for (int k = 0; k < oneRowData.size(); k++) {
                        row.createCell(k).setCellValue(oneRowData.get(k));
                    }
                }
            }
            writeWorkbook2File(fileName, workBook);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void writeMultiSheet(String fileName, List<List<List<String>>> dataList) {
        List<String> sheetName = new ArrayList<>();
        for (int i = 0; i < dataList.size(); i++) {
            sheetName.add("sheet" + i);
        }
        writeMultiSheet(fileName, sheetName, dataList);
    }

    private void writeMultiSheet(String fileName, List<String> sheetName, List<List<List<String>>> dataList) {
        if (isExcel2003(fileName)) {
            writeMultiSheet2003(fileName, sheetName, dataList);
            return;
        }
        writeMultiSheet2007(fileName, sheetName, dataList);
    }


    /**
     * main测试方法
     */
    public static void main(String[] args) {
        ExcelUtil poi = new ExcelUtil();
        String filePath = "test.xls";
        List<List<List<String>>> res = new ArrayList<>();
        List<List<String>> list = new ArrayList<>();
        List<String> addedList = new ArrayList<String>();
        addedList.add("我加了");
        addedList.add("这一行");
        addedList.add("的内容");
        list.add(addedList);
        //写一个工作簿，sheet名默认
        poi.write(filePath, list);
        //写一个工作簿，sheet名自己起
        poi.write(filePath, "sheetName", list);
        res.add(list);
        res.add(list);
        res.add(list);
        //写多个工作簿，sheet名默认
        poi.writeMultiSheet(filePath, res);
        List<String> sheetName = Arrays.asList("sheetName0", "sheetName1", "SheetName2");
        //写多个工作簿，sheet名自己起
        poi.writeMultiSheet(filePath,sheetName, res);
        System.out.println("已完成");
    }
}