package Util;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class WordtoPdfUtil {
     private static List<File> fileList = new ArrayList<>();

//     文件转换方法
    public void WordToPdf(String inputPeth){
        File inputWord = new File(inputPeth);
        File outputFile = new File(inputPeth.substring(0,inputPeth.indexOf("."))+".pdf");
        System.out.println(outputFile);
        try{
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("success");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //查找文件夹其子文件夹下的文件
    public static List<File> getFileList(String filePath) {
        File dir = new File(filePath);
        File[] files = dir.listFiles();
        if (files != null) {
            for (File file : files) {
                String fileName = file.getName();
                if (file.isDirectory()) {
                    getFileList(file.getAbsolutePath());
                } else if (fileName.endsWith("doc")||
                        fileName.endsWith("docx")) {
                    fileList.add(file);
                }
            }
        }
        return fileList;
    }



}
