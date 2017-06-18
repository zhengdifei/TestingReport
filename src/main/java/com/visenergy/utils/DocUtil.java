package com.visenergy.utils;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import sun.misc.BASE64Encoder;

import java.awt.*;
import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created  on 2017/6/8 0024.
 * @author zhengdifei
 * @desc word操作集合
 */
public class DocUtil {
    //后缀集合
    private static Map<String,Integer> SUFFIX_MAP = new HashMap<String, Integer>();

    static{
        //初始化后缀
        SUFFIX_MAP.put("doc", SaveFormat.DOC);
        SUFFIX_MAP.put("docx", SaveFormat.DOCX);
        SUFFIX_MAP.put("pdf", SaveFormat.PDF);

        //Aspose 加载license文件，没有license，导出文件有水印
        InputStream is = DocUtil.class.getClassLoader().getResourceAsStream("license.xml");
        License asposeLic = new License();
        try {
            asposeLic.setLicense(is);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * @desc 通过Aspose生成pdf文件
     * @param datas 模板数据，${key}格式
     * @param importFile 输入文件，doc格式
     * @param outputFile 输出文件
     */
    public static void createFileByAspose(Map<String,Object> datas,String importFile, String outputFile){
        try {
            //根据输入文件路径，构建document
            Document doc = new Document(importFile);
            //获取所有需要替换数据的key名称
            Iterator<String> keys = datas.keySet().iterator();
            while(keys.hasNext()){
                String key = keys.next();
                //当word中，需要替换是一个字符串，直接替换。替换格式${key}
                if( datas.get(key) instanceof String){
                    String value = String.valueOf(datas.get(key));

                    value = value.replace("\r\n","");

                    doc.getRange().replace("${" + key + "}",value,true,false);
                //当word中，需要替换的内容是一个Map对象。替换内容根据type关键字判断，可能是一个图片或者报表，
                }else if(datas.get(key) instanceof Map){
                    Map<String,Object> subObj = (Map<String,Object>) datas.get(key);
                    String type = String.valueOf(subObj.get("type"));
                    //当替换的内容是图片
                    if("image".equals(type)){
                        String url = String.valueOf(subObj.get("url"));
                        doc.getMailMerge().execute(new String[]{key},new Object[]{getImage(url)});
                    //替换的内容是饼图
                    }else if("pie".equals(type)){
                        //设置报表默认字体及颜色
                        ChartFactory.setChartTheme(getDefaultReportTheme());
                        //饼图数据集
                        DefaultPieDataset dpd = new DefaultPieDataset();
                        //构建饼图的值
                        Map<String,Number> pieMap = (Map<String,Number>)subObj.get("value");
                        for(Map.Entry<String,Number> pieEntry : pieMap.entrySet()){
                            dpd.setValue(pieEntry.getKey(),pieEntry.getValue());
                        }
                        //构建标题
                        String title = subObj.get("title")==null ? "":String.valueOf(subObj.get("title"));
                        //构建宽度
                        int width = subObj.get("width")==null ? 500:Integer.valueOf(String.valueOf(subObj.get("width")));
                        //构建高度
                        int height = subObj.get("height")==null ? 400:Integer.valueOf(String.valueOf(subObj.get("height")));
                        //可以查具体的API文档,第一个参数是标题，第二个参数是一个数据集，第三个参数表示是否显示Legend，第四个参数表示是否显示提示，第五个参数表示图中是否存在URL
                        JFreeChart chart = ChartFactory.createPieChart(title,dpd,true,true,false);

                        ByteArrayOutputStream baos = new ByteArrayOutputStream();

                        //使用一个面向application的工具类，将chart转换成JPEG格式的图片。第3个参数是宽度，第4个参数是高度。
                        ChartUtilities.writeChartAsJPEG(baos,chart,width,height);

                        doc.getMailMerge().execute(new String[]{key},new Object[]{baos.toByteArray()});
                    //替换的内容是柱状图，折线图，面积图
                    }else if("bar".equals(type) || "line".equals(type) || "area".equals(type)){
                        //设置报表默认字体及颜色
                        ChartFactory.setChartTheme(getDefaultReportTheme());
                        //构建标题
                        String title = subObj.get("title")==null ? "":String.valueOf(subObj.get("title"));
                        //构建x轴
                        String xAxis = subObj.get("xAxis")==null ? "":String.valueOf(subObj.get("xAxis"));
                        //构建y轴
                        String yAxis = subObj.get("yAxis")==null ? "":String.valueOf(subObj.get("yAxis"));
                        //构建宽度
                        int width = subObj.get("width")==null ? 500:Integer.valueOf(String.valueOf(subObj.get("width")));
                        //构建高度
                        int height = subObj.get("height")==null ? 400:Integer.valueOf(String.valueOf(subObj.get("height")));
                        // 创建数据源
                        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
                        //构建饼图的值
                        List<HashMap> barList = (List<HashMap>)subObj.get("value");
                        for (Map one: barList) {
                            Number value = one.get("value") == null ? 0:(Number) one.get("value");
                            String series = one.get("series") == null ? "":String.valueOf(one.get("series"));
                            String category = one.get("category") == null ? "":String.valueOf(one.get("category"));
                            // 放入数据
                            dataset.addValue(value,series,category);
                        }
                        JFreeChart chart = null;

                        if("bar".equals(type)){
                            chart = ChartFactory.createBarChart(title,xAxis,yAxis,
                                    dataset,// 数据
                                    PlotOrientation.VERTICAL,// 定位，VERTICAL：垂直
                                    true,// 是否显示图例注释(对于简单的柱状图必须是false)
                                    true,// 是否生成工具//没用过
                                    false);// 是否生成URL链接//没用过
                        }else if("line".equals(type)){

                            chart = ChartFactory.createLineChart(title,xAxis,yAxis,
                                    dataset,// 数据
                                    PlotOrientation.VERTICAL,// 定位，VERTICAL：垂直
                                    true,// 是否显示图例注释(对于简单的柱状图必须是false)
                                    true,// 是否生成工具//没用过
                                    false);// 是否生成URL链接//没用过
                        }else{
                            // 创建图形对象
                            chart = ChartFactory.createAreaChart(title,xAxis,yAxis,
                                    dataset, PlotOrientation.VERTICAL, true, true,
                                    false);
                            //设置前景色为50%透明
                            chart.getPlot().setForegroundAlpha(0.5F);
                        }

                        // 周围的背景色
                        chart.setBackgroundPaint(Color.white);
                        // 得到一个参考
                        CategoryPlot plot = (CategoryPlot) chart.getPlot();

                        //设置样式
                        Map styleMap = subObj.get("style") == null ? null : (Map)subObj.get("style");
                        if(styleMap != null){
                            String bcolorStr = styleMap.get("background") == null ? "#ffffff" : String.valueOf(styleMap.get("background"));
                            bcolorStr = bcolorStr.substring(1);
                            // 生成图片的背景色
                            plot.setBackgroundPaint(new Color(Integer.parseInt(bcolorStr,16)));
                        }
                        //字节输出流
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();

                        //使用一个面向application的工具类，将chart转换成JPEG格式的图片。第3个参数是宽度，第4个参数是高度。
                        ChartUtilities.writeChartAsJPEG(baos,chart,width,height);

                        doc.getMailMerge().execute(new String[]{key},new Object[]{baos.toByteArray()});
                    }else{
                        System.err.println("不支持的图标类型 ：" + type);
                    }
                }else if(datas.get(key) instanceof List){
                    List<Map<String,Object>> subObj = (List<Map<String,Object>>) datas.get(key);
                    //循环遍历
                    doc.getMailMerge().executeWithRegions(new MapMailMergeDataSource(subObj,key));
                }else{
                    System.err.println("不支持的数据结构，key ：" + key);
                }

            }
            //根据导出文件后缀，生产不同类型的文件，目前文件只支持doc,docx,pdf
            String suffix = outputFile.substring(outputFile.lastIndexOf(".") + 1);
           if(SUFFIX_MAP.get(suffix) != null){
               doc.save(outputFile,SUFFIX_MAP.get(suffix));
           }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static StandardChartTheme getDefaultReportTheme(){
        //创建主题样式
        StandardChartTheme sct = new StandardChartTheme("CN");
        //设置标题字体
        sct.setExtraLargeFont(new Font("隶书", Font.BOLD,20));
        //设置图例的字体
        sct.setRegularFont(new Font("宋书",Font.PLAIN,15));
        //设置轴向的字体
        sct.setLargeFont(new Font("宋书",Font.PLAIN,15));

        return sct;
    }
    /**
     * @desc 将一个图片生成base64数据
     * @param imgFile 图片文件路径
     * @return
     */
    public static String getImageStr(String imgFile){
        BASE64Encoder encoder = new BASE64Encoder();
        return encoder.encode(getImage(imgFile));
    }

    /**
     * @desc 将一个图片生成字节数组
     * @param imgFile 图片文件路径
     * @return
     */
    public static byte[] getImage(String imgFile){
        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(imgFile);
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }
}
