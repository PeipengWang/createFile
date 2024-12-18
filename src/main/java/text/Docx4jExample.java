package text;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;

public class Docx4jExample {
    public static void main(String[] args) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        ObjectFactory factory = new ObjectFactory();

        // 创建段落
        P paragraph = factory.createP();
        R run = factory.createR();
        Text text = factory.createText();
        text.setValue("This is a styled paragraph");

        run.getContent().add(text);
        paragraph.getContent().add(run);

        // 设置字体样式：加粗、斜体
        RPr rpr = factory.createRPr();
        BooleanDefaultTrue b = factory.createBooleanDefaultTrue();
        rpr.setB(b);
        BooleanDefaultTrue i = factory.createBooleanDefaultTrue();
        rpr.setI(i);
        run.setRPr(rpr);

        // 设置段落对齐
        PPr paragraphProperties = factory.createPPr();
        Jc alignment = factory.createJc();
        alignment.setVal(JcEnumeration.CENTER);  // 设置居中对齐
        paragraphProperties.setJc(alignment);
        paragraph.setPPr(paragraphProperties);

        // 将段落添加到文档
        wordMLPackage.getMainDocumentPart().addObject(paragraph);

        // 保存文件
        wordMLPackage.save(new File("styled_paragraph.docx"));
    }
}