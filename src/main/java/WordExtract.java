import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;
import java.util.*;
import java.util.regex.*;
import java.util.stream.Collectors;

/**
 * 本模块的主要功能：
 * 检查论文中的每一张图是否有对应的文字说明，
 * > 例如见图X.Y
 * 查论文中的每一张表是否有对应的文字说明，
 * > 例如见表X.Y
 * 返回所有没有文字说明的图和表列表
 * @author 谢亦敏
 * @version 1.0
 */
public class WordExtract {
    String originNamePattern = "^(%s\\s*(\\d+\\.?\\d*)\\s*[^\\s]*)(\\s*\\d+)?(?<!\\pP)$";
    String originInContentPattern ="%s\\s*(\\d+\\.?\\d*)([-到~～至—]%s?(\\d+\\.?\\d*))?";

    public static void main(String[] args) {
        if(args.length<1){
            System.out.println("请输入路径!");
        }
        try {
            ZipSecureFile.setMinInflateRatio(0);
            for(String url:args) {
                System.out.printf("------正在检查文件：%s------\n",url);
                new WordExtract().run(url);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 检查给定URL对应的docx中的图和表是否在正文中出现
     * @param url 给定URL
     * @throws IOException 如果指定文件不存在，则抛出IOException
     */
    public void run(String url) throws IOException{
        XWPFDocument doc = new XWPFDocument(new FileInputStream(url));

        System.out.println("------检查图片情况------");
        for(String i: checkType(doc,"图")){
            System.out.println(i);
        }

        System.out.println("------检查表格情况------");
        for(String i: checkType(doc,"表")){
            System.out.println(i);
        }
    }
    /**
     * 检查类型是否一一对应
     * @param docx 待检查文件
     * @return 所有错误的列表，格式如下：
     * 如果所查找图表在正文中不存在，则输出内容: "Type a.b 在正文中不存在"
     * 如果在正文中出现了目录里不存在的图表，则输出内容："正文中描述的Type a.b未在目录中出现"
     * 如果出现了如下情况：图1.3-1.1，输出内容："图片顺序错乱"
     */
    public List<String> checkType(XWPFDocument docx,String type){
        Map<String,Integer> checklist = getAllTypeNames(docx,type);
        Pattern skipPattern = Pattern.compile(String.format(originNamePattern,type));
        Pattern searchPattern = Pattern.compile(String.format(originInContentPattern,type,type));
        List<XWPFParagraph> allParagraphs = docx.getParagraphs();

        List<String> res = new ArrayList<>();

        for(int i=0;i<allParagraphs.size() && checklist.size()>0;i++){
            String curParagraph = allParagraphs.get(i).getText();
            if(curParagraph.startsWith(type)&&skipPattern.matcher(curParagraph).find())continue;

            Matcher matcher = searchPattern.matcher(curParagraph);
            while (matcher.find()){
                if(matcher.group(3)==null){
                    if(checklist.containsKey(matcher.group(1))){
                        checklist.put(matcher.group(1),-1);
                    }else{
                        res.add(String.format("正文中描述的%s%s未在目录中出现",type,matcher.group(1)));
                    }
                }else{
                    String[] am =matcher.group(1).split("\\.");
                    String[] xm=matcher.group(3).split("\\.");
                    if(am.length>1) {
                        //形式为图a.b-a.y或a.b-y
                        String a = am[0], b = am[1];
                        String x = xm[0];
                        if (!a.equals(x)) {
                            res.add(String.format("内容\"%s\"中，编号与所属章节不符",matcher.group(0)));
                            break;
                        }
                        if (xm.length > 1) x = xm[1];
                        int begin = Integer.parseInt(b), end = Integer.parseInt(x);
                        if (begin > end) {
                            res.add(String.format("内容\"%s\"中，编号顺序错乱",matcher.group(0)));
                            break;
                        }
                        for (int j = begin; j <= end; j++) {
                            //构建模式a.b
                            String id = a + "." + Integer.valueOf(j).toString();
                            if(checklist.containsKey(id)){
                                checklist.put(id,-1);
                            }else{
                                res.add(String.format("内容\"%s\"中，%s未在目录中出现",matcher.group(0),id));
                            }
                        }
                    }else{
                        //形式为图a-b
                        String a=am[0],b=xm[0];
                        int begin = Integer.parseInt(a), end = Integer.parseInt(b);
                        if (begin > end) {
                            res.add(String.format("内容\"%s\"中，编号顺序错乱",matcher.group(0)));
                            break;
                        }
                        for(int j=begin;j<=end;j++){
                            //构建模式a
                            String id=Integer.valueOf(j).toString();
                            checklist.remove(id);
                            if(checklist.containsKey(id)){
                                checklist.put(id,-1);
                            }else{
                                res.add(String.format("内容\"%s\"中，%s未在目录中出现",matcher.group(0),id));
                            }
                        }
                    }
                }
            }
        }
        checklist.entrySet()
                .stream()
                .filter(i -> i.getValue() != -1)
                .map(i -> String.format("%s%s未在正文中出现", type, i.getKey()))
                .forEach(res::add);
        return res;
    }
    /**
     * return all pictures' name from docx. If the document has a
     * TOC of pictures and named "图目录", that will be parsed as a returning,
     * or it will take all paragraphs which has such features as picture's name:
     * - start with "图 m.n"
     * - then not "所示"
     *
     * @param docx input document reference
     * @return all {"picture's id" = "pictures' paragraph id"}
     */
    public Map<String,Integer> getAllTypeNames(XWPFDocument docx,String type) {
        Map<String,Integer> result = new HashMap<>();
        List<XWPFParagraph> allParagraphs = docx.getParagraphs();
        Pattern regex = Pattern.compile(String.format(originNamePattern,type));
        //扫描全文检索图表名
        for(int i=0;i<allParagraphs.size();i++){
            String text = allParagraphs.get(i).getText();
            if(text.startsWith(type)){
                Matcher matcher = regex.matcher(text);
                if(matcher.find()){
                    result.put(matcher.group(2),i);
                }
            }
        }
        return result;
    }
}
