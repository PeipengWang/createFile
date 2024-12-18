package table;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;
import java.math.BigInteger;

public class Docx4jExample {
    public static void main(String[] args) throws Exception {
        // 创建 Word 文档
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        ObjectFactory factory = new ObjectFactory();

        // 创建表格
        Tbl table = factory.createTbl();
        TblPr tblPr = factory.createTblPr();
        TblBorders tblBorders = factory.createTblBorders();

        // 创建一个边框对象
        CTBorder border = factory.createCTBorder();
        border.setColor("000000");  // 黑色
        border.setSz(BigInteger.valueOf(4));  // 设置边框宽度
        border.setSpace(BigInteger.valueOf(0));  // 设置边框间距

        // 设置表格四个边框及内部边框
        tblBorders.setTop(border);
        tblBorders.setLeft(border);
        tblBorders.setBottom(border);
        tblBorders.setRight(border);
        tblBorders.setInsideH(border);
        tblBorders.setInsideV(border);

        // 将边框设置到表格属性中
        tblPr.setTblBorders(tblBorders);

        // 将表格属性设置到表格中
        table.setTblPr(tblPr);

        // 创建表格行和单元格
        Tr row = factory.createTr();
        Tc cell = factory.createTc();
        P paragraph = factory.createP();
        R run = factory.createR();
        Text text = factory.createText();
        text.setValue("Cell with Border");
        run.getContent().add(text);
        paragraph.getContent().add(run);
        cell.getContent().add(paragraph);
        row.getContent().add(cell);

        // 将行添加到表格
        table.getContent().add(row);

        // 将表格添加到文档
        wordMLPackage.getMainDocumentPart().addObject(table);

        // 保存文件
        wordMLPackage.save(new File("styled_table.docx"));
    }
}
