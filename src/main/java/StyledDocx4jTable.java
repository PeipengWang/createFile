import org.apache.commons.lang3.StringUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;

public class StyledDocx4jTable {
    private static String docxPath = "E:\\code\\creteFileDemo\\helloworld.docx";
    public static void main(String[] args) {
//        getWordprocessingMLPackage(docxPath);
        addParagraph();
    }
    static WordprocessingMLPackage getWordprocessingMLPackage(String docxPath) {
        WordprocessingMLPackage wordMLPackage = null;
        if(StringUtils.isEmpty(docxPath)) {
            try {
                wordMLPackage =  WordprocessingMLPackage.createPackage();
            } catch (InvalidFormatException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
        File file = new File(docxPath);
        if(file.isFile()) {
            try {
                wordMLPackage = WordprocessingMLPackage.load(file);
            } catch (Docx4JException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
        return wordMLPackage;
    }

    /**
     * 像文档中追加内容(默认支持中文)
     * 先清空,在生成,防重复
     */
    public static void addParagraph() {
        WordprocessingMLPackage wordprocessingMLPackage;
        try {
            //先加载word文档
            wordprocessingMLPackage = WordprocessingMLPackage.load(new File(docxPath));
//            Docx4J.load(new File(docxPath));

            //增加内容
            wordprocessingMLPackage.getMainDocumentPart().addParagraphOfText("你好!");
            wordprocessingMLPackage.getMainDocumentPart().addStyledParagraphOfText("Title", "这是标题!");
            wordprocessingMLPackage.getMainDocumentPart().addStyledParagraphOfText("Subtitle", " 这是二级标题!");

            wordprocessingMLPackage.getMainDocumentPart().addStyledParagraphOfText("Subject", "试一试");
            //保存文档
            wordprocessingMLPackage.save(new File(docxPath));
        } catch (Docx4JException e) {
           e.printStackTrace();
        }
    }




}