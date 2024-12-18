import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;

public class Docx4jExample {
    public static void main(String[] args) throws Exception {
        // 创建一个 Word 文档包
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

        // 创建一个段落
        ObjectFactory factory = new ObjectFactory();
        P paragraph = factory.createP();

        // 创建运行 (run)，即文本内容
        R run = factory.createR();
        Text text = factory.createText();
        text.setValue("Hello, docx4j!");
        run.getContent().add(text);

        // 将运行添加到段落中
        paragraph.getContent().add(run);

        // 将段落添加到文档
        wordMLPackage.getMainDocumentPart().addObject(paragraph);

        // 保存文档
        wordMLPackage.save(new File("example.docx"));
    }
}
