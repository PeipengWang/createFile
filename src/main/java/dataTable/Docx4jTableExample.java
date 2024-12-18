package dataTable;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;

public class Docx4jTableExample {
    public static void main(String[] args) throws Exception {
        // 创建 Word 文档
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

        // 创建一个标题
        ObjectFactory factory = Context.getWmlObjectFactory();
        P title = factory.createP();
        R titleRun = factory.createR();
        Text titleText = factory.createText();
        titleText.setValue("User Information Table");
        titleRun.getContent().add(titleText);
        title.getContent().add(titleRun);
        wordMLPackage.getMainDocumentPart().addObject(title);

        // 创建表格
        Tbl table = factory.createTbl();

        // 设置表格边框
        addBorders(table);

        // 添加表头
        addTableRow(factory, table, Arrays.asList("ID", "Name", "Age"));

        // 添加数据行
        addTableRow(factory, table, Arrays.asList("1", "Alice", "25"));
        addTableRow(factory, table, Arrays.asList("2", "Bob", "30"));
        addTableRow(factory, table, Arrays.asList("3", "Charlie", "28"));

        // 将表格添加到文档中
        wordMLPackage.getMainDocumentPart().addObject(table);

        // 保存文档
        wordMLPackage.save(new java.io.File("UserTable.docx"));
        System.out.println("Word document created: UserTable.docx");
    }

    // 添加表格行
    private static void addTableRow(ObjectFactory factory, Tbl table, List<String> cellValues) {
        Tr row = factory.createTr();
        for (String value : cellValues) {
            Tc cell = factory.createTc();

            // 创建段落并添加文本
            P paragraph = factory.createP();
            R run = factory.createR();
            Text text = factory.createText();
            text.setValue(value);

            run.getContent().add(text);
            paragraph.getContent().add(run);
            cell.getContent().add(paragraph);

            row.getContent().add(cell);
        }
        table.getContent().add(row);
    }

    // 设置表格边框
    private static void addBorders(Tbl table) {
        TblPr tblPr = Context.getWmlObjectFactory().createTblPr();
        TblBorders borders = Context.getWmlObjectFactory().createTblBorders();

        CTBorder border = Context.getWmlObjectFactory().createCTBorder();
        border.setColor("auto");
        border.setSz(new BigInteger("4"));
        border.setSpace(new BigInteger("10"));
        border.setVal(STBorder.SINGLE);

        borders.setTop(border);
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setInsideH(border);
        borders.setInsideV(border);

        tblPr.setTblBorders(borders);
        table.setTblPr(tblPr);
    }
}
