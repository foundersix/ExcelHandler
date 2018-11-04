package com.lkx;

import com.alibaba.excel.ExcelReader;

import java.io.*;
import java.net.URL;

public class App 
{
    public static void main( String[] args )
    {
        InputStream inputStream = null;
        File file = null;
        try {
            file = new File("C:\\Users\\lkx\\Desktop\\patient2.xlsx");

            String fileName = file.getName().substring(0,file.getName().lastIndexOf('.'));

            inputStream = new BufferedInputStream(new FileInputStream(file));
            // 解析每行结果在listener中处理
            ExcelListener listener = new ExcelListener(fileName);
            ExcelReader excelReader = new ExcelReader(inputStream, null, listener,true);
            excelReader.read();
        } catch (Exception e) {
              e.printStackTrace();
        } finally {
            try {
                if(inputStream != null)
                    inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
