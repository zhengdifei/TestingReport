package com.visenergy;

import com.visenergy.utils.DocUtil;
import net.sf.json.JSONObject;
import org.apache.commons.cli.*;

import java.io.*;
import java.net.URL;
import java.util.Map;

/**
 * Created by Administrator on 2017/6/7 0007.
 */
public class RunReport {
    public static void main(String[] args){
        //解析输入参数
        CommandLineParser parser = new BasicParser();
        Options options = new Options();
        options.addOption("h","help",false,"print this usage information");
        options.addOption("t","template file",true,"testing report template");
        options.addOption("o","output file path",true,"output file path");

        try {
            CommandLine cl = parser.parse(options,args);
            //模板文件路径
            String templateFile = "";
            //输出文件路径
            String outputFile = "";

            //帮助信息
            if(cl.hasOption('h')){
                System.out.println("-t  tesing report template[requied]");
                System.out.println("-o  output file path[requied]");
                System.out.println("-h  help infomation");
                System.exit(0);
            }
            //输入模板文件，输出文件处理及异常处理
            boolean lackParam = false;
            if(cl.hasOption('t')){
                templateFile = cl.getOptionValue('t');
            }else {
                System.err.println("The -t parameter is missing!");
                lackParam = true;
            }

            if(cl.hasOption('o')){
                outputFile = cl.getOptionValue('o');
            }else {
                System.err.println("The -o parameter is missing!");
                lackParam = true;
            }
            //异常信息提示
            if(lackParam) {
                System.err.println("The parameter info:");
                System.err.println("    -t  tesing report template[requied]");
                System.err.println("    -o  output file path[requied]");
                System.err.println("    -h  help infomation");
                System.exit(0);
            }

            System.out.println("测试报告模板文件位置：" + templateFile);
            System.out.println("生成文件位置：" + outputFile);

            //输入数据流，输入数据必须是json格式数据
            //读取资源文件或者jar包中data.json文件
            InputStream is = null;//RunReport.class.getClassLoader().getResourceAsStream("data.json");
            //InputStream is = RunReport.class.getClassLoader().getResourceAsStream("data.json");
            if(is == null){
                is = new FileInputStream("data.json");
            }
            BufferedReader br = new BufferedReader(new InputStreamReader(is,"utf-8"));
            StringBuffer dataStr = new StringBuffer();
            String line = null;
            while ((line = br.readLine()) != null){
                dataStr.append(line);
            }

            if(dataStr.length() > 0) {
                JSONObject dataJson = JSONObject.fromObject(dataStr.toString());
                Map<String,Object> paramMap = (Map<String, Object>) dataJson;

                //数据解析
                DocUtil.createFileByAspose(paramMap,templateFile,outputFile);
            }else{
                System.err.println("输入数据流无数据");
            }

        } catch (ParseException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
