package header;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;

/**
 * 添加标题
 */
public class Docx4jExample {
    public static void main(String[] args) throws Exception {
        // 创建 Word 文档
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        ObjectFactory factory = new ObjectFactory();

        // 创建段落
        P titleParagraph = factory.createP();
        R titleRun = factory.createR();
        Text titleText = factory.createText();
        titleText.setValue("Document Title");

        titleRun.getContent().add(titleText);
        titleParagraph.getContent().add(titleRun);

        // 创建段落属性对象（PPr）
        PPr titlePPr = factory.createPPr();

        // 设置段落样式为 Heading 1
        PPrBase.PStyle titlePStyle = new PPrBase.PStyle();
        titlePStyle.setVal("Heading1");  // 设置样式为 Heading1
        titlePPr.setPStyle(titlePStyle);

        // 设置段落对齐方式为居中
        Jc titleAlignment = factory.createJc();
        titleAlignment.setVal(JcEnumeration.CENTER);  // 设置居中对齐
        titlePPr.setJc(titleAlignment);

        // 将段落属性应用到段落
        titleParagraph.setPPr(titlePPr);

        // 将段落添加到文档
        wordMLPackage.getMainDocumentPart().addObject(titleParagraph);

        // 保存文件
        wordMLPackage.save(new File("example_with_heading1.docx"));
    }
}
