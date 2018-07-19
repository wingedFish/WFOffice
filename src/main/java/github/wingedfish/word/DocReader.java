package github.wingedfish.word;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

/**
 * wingedfish-wfoffice
 *
 * @author lixiuhai
 * @create 2018/07/17 09:50
 **/
public class DocReader {

    public static void main(String[] args) {
        String dirPath1 = "D:\\应用平台研发组工作\\问题梳理\\2018Q1工作总结";
        String dirPath2 = "D:\\应用平台研发组工作\\问题梳理\\2018Q2工作总结";

        DocReader docReader = new DocReader();

        Map<String, String> idea = new HashMap<>();
//        idea.putAll(docReader.readDir(dirPath1));
        idea.putAll(docReader.readDir(dirPath2));


        docReader.writeCSV(idea);

    }

    public Map<String,String> readDir(String dirPath){
        File dir = new File(dirPath);
        File[] files = dir.listFiles();


        Map<String, String> idea = new HashMap<>();
        for (int i = 0; i < files.length; i++) {

            Map content = read(files[i]);
            idea.putAll(content);
        }

        return idea;
    }

    public Map read(File file) {
        try {
            XWPFDocument xdoc = new XWPFDocument(new FileInputStream(file));
            XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
            String text = extractor.getText();
            return resolveDocument(text, file.getName());

        } catch (Exception e) {
            System.err.println("======" + file.getName() + e);
        }
        return null;
    }

    private Map resolveDocument(String text, String fileName) {
        Map<String, String> content = new HashMap<>();

        int index = text.indexOf("问题和建议");
        if (index < 0) {
            throw new RuntimeException("No found the key !");
        }

        int nameIndex = text.indexOf("姓名");
        if (nameIndex < 0) {
            content.put("问题描述", fileName);
//            throw new RuntimeException("No found the key-name !");
        }

        String name = text.substring(nameIndex + 3, nameIndex + 6).replace("\n","");
        String idea = text.substring(index).replace("\n","");

        content.put(name, "\"" + idea + "\"");

//        System.out.println(idea);
//        System.out.println(name);
//        System.out.println("==================");
        return content;
    }


    private void writeCSV(Map<String, String> content) {
        Path path = Paths.get("D:\\应用平台研发组工作\\问题梳理\\问题总结Q2.csv");
        try (BufferedWriter writer = Files.newBufferedWriter(path)) {
            writer.write("序号,提出时间,问题描述,提出人,解决方案,备注");
            writer.newLine();
            Set<Map.Entry<String, String>> entrySet = content.entrySet();
            int i = 1;
            for (Iterator<Map.Entry<String, String>> iterator = entrySet.iterator(); iterator.hasNext(); ) {
                Map.Entry<String, String> entry = iterator.next();
                String key = entry.getKey();
                String value = entry.getValue();
                String info = i + ",2018Q2," + value + "," + key + ",无,无";
                writer.write(info);
                i++;
                writer.newLine();
                System.out.println("content = [" + info + "]");
            }
            writer.flush();
            writer.close();

        } catch (Exception e) {


        }

    }
}
