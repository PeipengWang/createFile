import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;

public class zhenghe {
    public static void main(String[] args) throws Docx4JException {
        // 创建 Word 文档
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        addHead(wordMLPackage);
        addtable(wordMLPackage);
        // 保存文件
        wordMLPackage.save(new File("aaa.docx"));
    }
    public static WordprocessingMLPackage addHead(WordprocessingMLPackage wordMLPackage){
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
        return wordMLPackage;
    }
    public static WordprocessingMLPackage addtable(WordprocessingMLPackage wordMLPackage) throws Docx4JException {
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
        addTableRowWithShading(factory, table, Arrays.asList("ID", "Name", "Age"), "D3D3D3");

        // 添加数据行
        addTableRow(factory, table, Arrays.asList("1", "Alice", "25"));
        addTableRow(factory, table, Arrays.asList("2", "Bob", "30"));
        addTableRow(factory, table, Arrays.asList("3", "Charlie", "28"));

        // 将表格添加到文档中
        wordMLPackage.getMainDocumentPart().addObject(table);

        return wordMLPackage;
    }
    // 添加带背景的表头行
    private static void addTableRowWithShading(ObjectFactory factory, Tbl table, List<String> rowData, String hexColor) {
        Tr row = factory.createTr();
        for (String cellData : rowData) {
            Tc cell = factory.createTc();
            P cellParagraph = factory.createP();
            R cellRun = factory.createR();
            Text cellText = factory.createText();
            cellText.setValue(cellData);
            cellRun.getContent().add(cellText);
            cellParagraph.getContent().add(cellRun);
            cell.getContent().add(cellParagraph);

            // 设置背景色
            TcPr tcPr = factory.createTcPr();
            CTShd shading = new CTShd();
            shading.setVal(STShd.CLEAR);  // 清除已有阴影
            shading.setColor("auto");
            shading.setFill(hexColor);   // 设置背景颜色 (灰色)
            tcPr.setShd(shading);
            cell.setTcPr(tcPr);

            row.getContent().add(cell);
        }
        table.getContent().add(row);
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
