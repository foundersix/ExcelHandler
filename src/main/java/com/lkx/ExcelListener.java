package com.lkx;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;


import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

public class ExcelListener extends AnalysisEventListener<Object> {

    private long start = 0;
    private String fileName;
    //自定义用于暂时存储data。
    //可以通过实例获取该值
    private List<Object> datas = new ArrayList<Object>();
    private StringBuffer sb = new StringBuffer(2000);
    private long flag = 0;


    public ExcelListener(String fileName){
        this.start = System.currentTimeMillis();
        this.fileName = fileName;
    }

    @Override
    public void invoke(Object object, AnalysisContext context) {


        //System.out.println("当前行："+context.getCurrentRowNum());
        //System.out.println(object);

        String currRow1 = object.toString().replace(", ",",");
        String currRow2 = currRow1.substring(1,currRow1.length()-15)+"\n";
        sb.append(currRow2);
        if((context.getCurrentRowNum()+1) % 1536 == 0){
            String result = sb.toString();
            this.flag++;
            doSave(result);//根据自己业务做处理
            sb.delete(0,sb.length());
        }
    }

    /**
     * "获取类路径需要改善"
     * @param result
     */
    private void doSave(String result) {

        URL base = this.getClass().getResource("");
        File file = null;
        try{
            String saveDir = new File(base.getFile(),
                    "../../../../src/resource/"+fileName).getCanonicalPath();
            File saveDirectory = new File(saveDir);
            if(!saveDirectory.exists()){
                saveDirectory.mkdir();
            }
            String savePath = saveDirectory.getCanonicalPath() + "\\"+fileName+ "_"+ this.flag +".csv";
            file = new File(savePath);
            if(!file.exists()){
                file.createNewFile();
            }
            BufferedWriter out = new BufferedWriter(new FileWriter(file));
            out.write(result);
            out.close();
        }catch(IOException e){
            e.printStackTrace();
        }



    }
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        if(this.sb.length() > 0){
            String result = sb.toString();
            this.flag++;
            doSave(result);
            sb.delete(0,sb.length());
        }

        System.out.print((System.currentTimeMillis() - start )/1000+"秒");
    }
/*    public List<Object> getDatas() {
        return datas;
    }
    public void setDatas(List<Object> datas) {
        this.datas = datas;
    }*/
}
