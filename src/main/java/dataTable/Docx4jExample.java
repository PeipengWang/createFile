package dataTable;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;
import java.math.BigInteger;

public class Docx4jExample {
    public static void main(String[] args) throws Exception {
        // 创建 Word 文档
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        ObjectFactory factory = new ObjectFactory();

        // 创建一个表格
        Tbl table = factory.createTbl();

        // 创建表格属性并设置
        TblPr tblPr = factory.createTblPr();

        // 设置表格宽度
        TblWidth tblWidth = factory.createTblWidth();
        tblWidth.setW(BigInteger.valueOf(5000));  // 设置表格宽度为5000单位
        tblPr.setTblW(tblWidth);

        // 设置表格边框
        TblBorders tblBorders = factory.createTblBorders();

        // 创建边框
        CTBorder border = factory.createCTBorder();
        border.setVal(STBorder.SINGLE);  // 设置边框样式为单线
        border.setSz(BigInteger.valueOf(4));  // 设置边框宽度
        border.setSpace(BigInteger.valueOf(0));  // 设置边框间距

        // 为表格的各个边框设置样式
        tblBorders.setTop(border);
        tblBorders.setLeft(border);
        tblBorders.setBottom(border);
        tblBorders.setRight(border);
        tblBorders.setInsideH(border);
        tblBorders.setInsideV(border);

        tblPr.setTblBorders(tblBorders);

        // 设置表格头部样式（例如，第一行加粗）
        Tr row = factory.createTr();
        Tc headerCell = factory.createTc();
        P headerParagraph = factory.createP();
        R headerRun = factory.createR();
        Text headerText = factory.createText();
        headerText.setValue("Header 1");
        headerRun.getContent().add(headerText);
        headerParagraph.getContent().add(headerRun);
        headerCell.getContent().add(headerParagraph);
        row.getContent().add(headerCell);

        Tc headerCell2 = factory.createTc();
        P headerParagraph2 = factory.createP();
        R headerRun2 = factory.createR();
        Text headerText2 = factory.createText();
        headerText2.setValue("Header 2");
        headerRun2.getContent().add(headerText2);
        headerParagraph2.getContent().add(headerRun2);
        headerCell2.getContent().add(headerParagraph2);
        row.getContent().add(headerCell2);

        // 将表头添加到表格
        table.getContent().add(row);

        // 创建表格数据行
        for (int i = 0; i < 5; i++) {
            Tr dataRow = factory.createTr();

            Tc cell = factory.createTc();
            P cellParagraph = factory.createP();
            R cellRun = factory.createR();
            Text cellText = factory.createText();
            cellText.setValue("Row " + (i + 1) + " Cell 1");
            cellRun.getContent().add(cellText);
            cellParagraph.getContent().add(cellRun);
            cell.getContent().add(cellParagraph);
            dataRow.getContent().add(cell);

            Tc cell2 = factory.createTc();
            P cellParagraph2 = factory.createP();
            R cellRun2 = factory.createR();
            Text cellText2 = factory.createText();
            cellText2.setValue("Row " + (i + 1) + " Cell 2");
            cellRun2.getContent().add(cellText2);
            cellParagraph2.getContent().add(cellRun2);
            cell2.getContent().add(cellParagraph2);
            dataRow.getContent().add(cell2);

            // 将数据行添加到表格
            table.getContent().add(dataRow);
        }

        // 将表格添加到文档
        wordMLPackage.getMainDocumentPart().addObject(table);

        // 保存文件
        wordMLPackage.save(new File("example_with_styled_table.docx"));
    }
}
