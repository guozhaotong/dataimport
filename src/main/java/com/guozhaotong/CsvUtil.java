package com.guozhaotong;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class CsvUtil {
    public List<List<String>> read(String filePath) {
        File csv = new File(filePath);  // CSV文件路径
        BufferedReader br = null;
        try {
            br = new BufferedReader(new FileReader(csv));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        List<List<String>> dataList = new ArrayList<>();
        try {
            String line = "";
            while ((line = br.readLine()) != null)  //读取到的内容给line变量
            {
                List<String> everyLine = Arrays.asList(line.split(","));
                dataList.add(everyLine);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return dataList;
    }

    public void write(String filePath, List<List<String>> dataList) {
        FileOutputStream out = null;
        OutputStreamWriter osw = null;
        BufferedWriter bw = null;
        try {
            File finalCSVFile = new File(filePath);
            out = new FileOutputStream(finalCSVFile);
            osw = new OutputStreamWriter(out, "UTF-8");
            // 手动加上BOM标识
            osw.write(new String(new byte[]{(byte) 0xEF, (byte) 0xBB, (byte) 0xBF}));
            bw = new BufferedWriter(osw);
            /**
             * 往CSV中写新数据
             */
            if (dataList != null && !dataList.isEmpty()) {
                for (List<String> line : dataList) {
                    for (int i = 0; i < line.size() - 1; i++) {
                        bw.append(line.get(i).toString().trim());
                        bw.append(",");

                    }
                    bw.append(line.get(line.size() - 1).toString().trim());
                    bw.append("\r");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.out.println("已完成");
            try {
                if (bw != null) {
                    bw.close();
                    bw = null;
                }
                if (osw != null) {
                    osw.close();
                    osw = null;
                }
                if (out != null) {
                    out.close();
                    out = null;
                }
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
    }

    public static void main(String[] args){
        CsvUtil csvUtil = new CsvUtil();
        List<List<String>> dataList = csvUtil.read("test.csv");
        System.out.println(dataList);
    }
}