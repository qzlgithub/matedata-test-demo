package net.matedata.utils;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.ExceptionConverter;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPTableEvent;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class PDFTest
{

    public static void main(String[] args) throws Exception
    {
        ZipOutputStream zip = new ZipOutputStream(new FileOutputStream("C:\\Users\\JinKX\\Desktop\\" + "zipPDF.zip"));
        for(int j = 1; j <= 10; j++)
        {
            ZipEntry entry = new ZipEntry("hello_" + j + ".pdf");

            zip.putNextEntry(entry);

            Document document = new Document(PageSize.A4, 0, 0, 50, 0);
            /*String pdfName = "HelloWorld.pdf";
            PdfWriter writer = PdfWriter.getInstance(document,new FileOutputStream("C:\\Users\\JinKX\\Desktop\\"+pdfName));*/

            PdfWriter writer = PdfWriter.getInstance(document, zip);
            writer.setCloseStream(false);

            //设置页面布局
            writer.setViewerPreferences(PdfWriter.PageLayoutOneColumn);

            //为这篇文档设置页面事件(X/Y)
            //writer.setPageEvent(new PageXofYTest());

            //创建BaseFont对象,指明字体,编码方式,是否嵌入
            BaseFont bf = BaseFont.createFont("C:\\Windows\\Fonts\\simkai.ttf", BaseFont.IDENTITY_H, false);

            //创建Font对象,将基础字体对象,字体大小
            Font font0 = new Font(bf, 8, Font.NORMAL);
            Font font = new Font(bf, 13, Font.NORMAL);
            Font font1 = new Font(bf, 15, Font.BOLD);
            Font font2 = new Font(bf, 20, Font.BOLD);
            Font font3 = new Font(bf, 22, Font.NORMAL);

            //隔行换色事件
            PdfPTableEvent event = new AlternatingBackground();

            //table1
            PdfPTable table1 = new PdfPTable(3);
            table1.setTotalWidth(new float[]{100f, 115f, 100f});
            table1.setSpacingBefore(50f);
            PdfPCell cell1 = new PdfPCell(new Paragraph("", font));
            cell1.setBorder(0);
            PdfPCell cellx2 = new PdfPCell(new Paragraph("本报告仅限于您在已授权情况下了解授权人信用状况下使用", font));
            PdfPCell cellx3 = new PdfPCell();
            cellx3.setBorder(0);
            String imagePath1 = "D:\\yuan.png";
            Image image1 = Image.getInstance(imagePath1);
            image1.setWidthPercentage(70);
            cellx3.addElement(image1);
            table1.addCell(cellx3);
            table1.addCell(cell1);
            table1.addCell(cellx2);

            //table2
            PdfPTable table2 = new PdfPTable(4);
            table2.setTotalWidth(new float[]{200f, 200f, 200f, 200f});
            table2.setSpacingAfter(30f);
            PdfPCell cell2 = new PdfPCell();
            cell2 = mergeCol("报告编号:20392323", font, 4);
            cell2.setBorder(0);
            cell2.setHorizontalAlignment(Element.ALIGN_LEFT);
            table2.addCell(cell2);
            table2.addCell(getPDFCell("被查询人", font1));
            table2.addCell(getPDFCell("身份证号码", font1));
            table2.addCell(getPDFCell("手机号码", font1));
            table2.addCell(getPDFCell("查询操作人", font1));

            table2.addCell(getPDFCell("王尼玛", font));
            table2.addCell(getPDFCell("**...**", font));
            table2.addCell(getPDFCell("183****900", font));
            table2.addCell(getPDFCell("王某某", font));

            //table3
            PdfPTable table3 = new PdfPTable(4);
            table3.setTotalWidth(new float[]{200f, 200f, 200f, 200f});
            table3.setSpacingAfter(30f);
            PdfPCell cell3 = new PdfPCell();
            PdfPCell cell4 = new PdfPCell();
            cell4 = mergeCol("--个人基本信息--", font, 4);
            cell4.setBorder(0);
            cell3 = mergeCol("身份信息", font1, 4);
            table3.addCell(cell4);
            table3.addCell(cell3);
            table3.addCell(getPDFCell("姓名:", font1));
            table3.addCell(getPDFCell("王尼玛", font));
            table3.addCell(getPDFCell("民族:", font1));
            table3.addCell(getPDFCell("汉", font));

            table3.addCell(getPDFCell("性别:", font1));
            table3.addCell(getPDFCell("男", font));
            table3.addCell(getPDFCell("有效期:", font1));
            table3.addCell(getPDFCell("2008/12--2018/12", font));

            table3.addCell(getPDFCell("户口所在地:", font1));
            table3.addCell(getPDFCell("新疆", font));
            table3.addCell(getPDFCell("现居地址:", font1));
            table3.addCell(getPDFCell("天津新港", font));

            //table4
            PdfPTable table4 = new PdfPTable(4);
            table4.setTotalWidth(new float[]{200f, 200f, 200f, 200f});
            table4.setSpacingAfter(30f);
            PdfPCell cell5 = new PdfPCell();
            cell5 = mergeCol("工作信息", font1, 4);
            table4.addCell(cell5);
            table4.addCell(getPDFCell("时间", font1));
            table4.addCell(getPDFCell("单位名称", font1));
            table4.addCell(getPDFCell("职务", font1));
            table4.addCell(getPDFCell("单位地址", font1));

            table4.addCell(getPDFCell("2017/06--2018/03", font));
            table4.addCell(getPDFCell("知名单位", font));
            table4.addCell(getPDFCell("保洁", font));
            table4.addCell(getPDFCell("松花江路109号", font));

            table4.addCell(getPDFCell("2017/06--2018/03", font));
            table4.addCell(getPDFCell("知名单位", font));
            table4.addCell(getPDFCell("保洁", font));
            table4.addCell(getPDFCell("松花江路109号", font));

            //table5
            PdfPTable table5 = new PdfPTable(3);
            table5.setTotalWidth(new float[]{200f, 200f, 200f});
            table5.setSpacingAfter(30f);
            PdfPCell cell6 = new PdfPCell();
            cell6 = mergeCol("婚姻情况", font1, 3);
            table5.addCell(cell6);
            table5.addCell(getPDFCell("婚姻状态", font1));
            table5.addCell(getPDFCell("配偶姓名", font1));
            table5.addCell(getPDFCell("配偶身份证号", font1));

            table5.addCell(getPDFCell("已婚", font));
            table5.addCell(getPDFCell("王大仙", font));
            table5.addCell(getPDFCell("12321*****2312", font));

            //table6
            PdfPTable table6 = new PdfPTable(2);
            table6.setTotalWidth(new float[]{200f, 200f});
            table6.setSpacingAfter(30f);
            PdfPCell cell7 = new PdfPCell();
            cell7 = mergeCol("家庭成员", font1, 2);
            table6.addCell(cell7);
            table6.addCell(getPDFCell("姓名", font1));
            table6.addCell(getPDFCell("关系", font1));

            table6.addCell(getPDFCell("王大拿", font));
            table6.addCell(getPDFCell("父亲", font));

            table6.addCell(getPDFCell("王大仙", font));
            table6.addCell(getPDFCell("妻子", font));

            //table7
            PdfPTable table7 = new PdfPTable(3);
            table7.setTotalWidth(new float[]{200f, 200f, 200f});
            table7.setSpacingAfter(15f);
            PdfPCell cell8 = new PdfPCell();
            cell8 = mergeCol("婚姻情况", font1, 3);
            table7.addCell(cell8);
            table7.addCell(getPDFCell("品牌", font1));
            table7.addCell(getPDFCell("估价", font1));
            table7.addCell(getPDFCell("车牌归属地", font1));

            table7.addCell(getPDFCell("克尔维特", font));
            table7.addCell(getPDFCell("280W", font));
            table7.addCell(getPDFCell("北京", font));

            //table8
            PdfPTable table8 = new PdfPTable(2);
            table8.setTotalWidth(new float[]{200f, 200f});
            table8.setSpacingAfter(15f);
            PdfPCell cell9 = new PdfPCell();
            cell9 = mergeCol("说明", font, 2);
            cell9.setBorder(0);
            cell9.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell10 = new PdfPCell();
            cell10 = mergeCol("1、以上数据均来自对数据分析得出，仅为业务层面提供参考;", font, 2);
            cell10.setBorder(0);
            cell10.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell11 = new PdfPCell();
            cell11 = mergeCol("2、个人信息十分重要，请不要泄露;", font, 2);
            cell11.setBorder(0);
            cell11.setHorizontalAlignment(Element.ALIGN_LEFT);
            table8.addCell(cell9);
            table8.addCell(cell10);
            table8.addCell(cell11);

            /*
             *第3页
             * table9
             * */
            PdfPTable table9 = new PdfPTable(2);
            table9.setTotalWidth(new float[]{200f, 200f});
            table9.setSpacingAfter(10f);
            PdfPCell cell12 = new PdfPCell();
            cell12 = mergeCol("查询人:张某", font1, 2);
            cell12.setBorder(0);
            cell12.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell13 = new PdfPCell();
            cell13 = mergeCol("编号:xhsgs45", font1, 1);
            cell13.setBorder(0);
            cell13.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell14 = new PdfPCell();
            cell14 = mergeCol("2018/02/23 12:21:23", font1, 1);
            cell14.setBorder(0);
            cell14.setHorizontalAlignment(Element.ALIGN_RIGHT);

            table9.addCell(cell12);
            table9.addCell(cell13);
            table9.addCell(cell14);

            table9.addCell(getPDFCell("历史单日最多注册平台", font));
            table9.addCell(getPDFCell("", font));

            table9.addCell(getPDFCell("", font));
            table9.addCell(getPDFCell("", font));

            //table10
            PdfPTable table10 = new PdfPTable(10);
            table10.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f});
            table10.setSpacingAfter(10f);
            table10.addCell(getPDFCell("注册最早出现时间", font));
            table10.addCell(getPDFCell("注册最近出现时间", font));
            table10.addCell(getPDFCell("历史注册平台总数", font));
            table10.addCell(getPDFCell("今日注册平台个数", font));
            table10.addCell(getPDFCell("近3天注册平台个数", font));
            table10.addCell(getPDFCell("近7天注册平台个数", font));
            table10.addCell(getPDFCell("近15天注册平台个数", font));
            table10.addCell(getPDFCell("近30天注册平台个数", font));
            table10.addCell(getPDFCell("近60天注册平台个数", font));
            table10.addCell(getPDFCell("近90天注册平台个数", font));

            for(int i = 0; i < 10; i++)
            {
                table10.addCell(getPDFCell("", font));
            }

            //table11
            PdfPTable table11 = new PdfPTable(4);
            table11.setTotalWidth(new float[]{200f, 200f, 200f, 200f});
            table11.setSpacingAfter(15f);
            table11.getDefaultCell().setBorder(0);
            table11.addCell(getPDFCell("序号", font));
            table11.addCell(getPDFCell("注册平台编号", font));
            table11.addCell(getPDFCell("平台类型", font));
            table11.addCell(getPDFCell("最早注册时间", font));

            table11.addCell(getPDFCell("1", font));
            table11.addCell(getPDFCell("Ssw564", font));
            table11.addCell(getPDFCell("金融", font));
            table11.addCell(getPDFCell("2015/06/02 15:45:12", font));

            table11.addCell(getPDFCell("2", font));
            table11.addCell(getPDFCell("SD", font));
            table11.addCell(getPDFCell("股票", font));
            table11.addCell(getPDFCell("2018/02/23 16:12:23", font));

            //table12
            PdfPTable table12 = new PdfPTable(2);
            table12.setTotalWidth(new float[]{200f, 200f});
            table12.setSpacingAfter(15f);
            PdfPCell cell15 = new PdfPCell();
            cell15 = mergeCol("说明", font, 2);
            cell15.setBorder(0);
            cell15.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell16 = new PdfPCell();
            cell16 = mergeCol("1、该报告主要展示被查询人在多个借款平台的注册情况;", font, 2);
            cell16.setBorder(0);
            cell16.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell17 = new PdfPCell();
            cell17 = mergeCol("2、该结果均来自对自有数据源的分析，其结果仅对本数据源负责;", font, 2);
            cell17.setBorder(0);
            cell17.setHorizontalAlignment(Element.ALIGN_LEFT);
            table12.addCell(cell15);
            table12.addCell(cell16);
            table12.addCell(cell17);

            /*
             *第4页
             *table13
             * */
            PdfPTable table13 = new PdfPTable(2);
            table13.setTotalWidth(new float[]{200f, 200f});
            table13.setSpacingAfter(10f);
            PdfPCell cell18 = new PdfPCell();
            cell18 = mergeCol("查询人:张某", font1, 2);
            cell18.setBorder(0);
            cell18.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell19 = new PdfPCell();
            cell19 = mergeCol("编号:xhsgs45", font1, 1);
            cell19.setBorder(0);
            cell19.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell20 = new PdfPCell();
            cell20 = mergeCol("2018/02/23 12:21:23", font1, 1);
            cell20.setBorder(0);
            cell20.setHorizontalAlignment(Element.ALIGN_RIGHT);

            table13.addCell(cell18);
            table13.addCell(cell19);
            table13.addCell(cell20);

            table13.addCell(getPDFCell("历史最大放款金额", font));
            table13.addCell(getPDFCell("", font));

            table13.addCell(getPDFCell("", font));
            table13.addCell(getPDFCell("", font));

            //table14
            PdfPTable table14 = new PdfPTable(10);
            table14.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f});
            table14.setSpacingAfter(10f);
            table14.addCell(getPDFCell("放款最早出现时间", font));
            table14.addCell(getPDFCell("放款最近出现时间", font));
            table14.addCell(getPDFCell("历史放款平台总数", font));
            table14.addCell(getPDFCell("今日放款平台个数", font));
            table14.addCell(getPDFCell("近3天放款平台个数", font));
            table14.addCell(getPDFCell("近7天放款平台个数", font));
            table14.addCell(getPDFCell("近15天放款平台个数", font));
            table14.addCell(getPDFCell("近30天放款平台个数", font));
            table14.addCell(getPDFCell("近60天放款平台个数", font));
            table14.addCell(getPDFCell("近90天放款平台个数", font));

            for(int i = 0; i < 10; i++)
            {
                table14.addCell(getPDFCell("", font));
            }

            //table15
            PdfPTable table15 = new PdfPTable(5);
            table15.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f});
            table15.setSpacingAfter(15f);
            table15.getDefaultCell().setBorder(0);
            table15.addCell(getPDFCell("序号", font));
            table15.addCell(getPDFCell("放款平台编号", font));
            table15.addCell(getPDFCell("平台类型", font));
            table15.addCell(getPDFCell("最早放款时间", font));
            table15.addCell(getPDFCell("最近放款时间", font));

            table15.addCell(getPDFCell("1", font));
            table15.addCell(getPDFCell("Ssw564", font));
            table15.addCell(getPDFCell("金融", font));
            table15.addCell(getPDFCell("2015/06/02 15:45:12", font));
            table15.addCell(getPDFCell("2015/06/02 15:45:12", font));

            table15.addCell(getPDFCell("2", font));
            table15.addCell(getPDFCell("SD", font));
            table15.addCell(getPDFCell("股票", font));
            table15.addCell(getPDFCell("2018/02/23 16:12:23", font));
            table15.addCell(getPDFCell("2015/06/02 15:45:12", font));

            //table16
            PdfPTable table16 = new PdfPTable(2);
            table16.setTotalWidth(new float[]{200f, 200f});
            table16.setSpacingAfter(15f);
            PdfPCell cell21 = new PdfPCell();
            cell21 = mergeCol("说明", font, 2);
            cell21.setBorder(0);
            cell21.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell22 = new PdfPCell();
            cell22 = mergeCol("1、该报告展示被查询人在其他平台通过借款情况;", font, 2);
            cell22.setBorder(0);
            cell22.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell23 = new PdfPCell();
            cell23 = mergeCol("2、该报告仅为自有数据统计;", font, 2);
            cell23.setBorder(0);
            cell23.setHorizontalAlignment(Element.ALIGN_LEFT);
            table16.addCell(cell21);
            table16.addCell(cell22);
            table16.addCell(cell23);

            /*
             *第5页
             *table17
             * */
            PdfPTable table17 = new PdfPTable(2);
            table17.setTotalWidth(new float[]{200f, 200f});
            table17.setSpacingAfter(10f);
            PdfPCell cell24 = new PdfPCell();
            cell24 = mergeCol("查询人:张某", font1, 2);
            cell24.setBorder(0);
            cell24.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell25 = new PdfPCell();
            cell25 = mergeCol("编号:xhsgs45", font1, 1);
            cell25.setBorder(0);
            cell25.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell26 = new PdfPCell();
            cell26 = mergeCol("2018/02/23 12:21:23", font1, 1);
            cell26.setBorder(0);
            cell26.setHorizontalAlignment(Element.ALIGN_RIGHT);

            table17.addCell(cell24);
            table17.addCell(cell25);
            table17.addCell(cell26);

            table17.addCell(getPDFCell("历史最大拒贷金额", font));
            table17.addCell(getPDFCell("", font));

            table17.addCell(getPDFCell("", font));
            table17.addCell(getPDFCell("", font));

            //table18
            PdfPTable table18 = new PdfPTable(10);
            table18.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f});
            table18.setSpacingAfter(10f);
            table18.addCell(getPDFCell("拒贷最早出现时间", font));
            table18.addCell(getPDFCell("拒贷最近出现时间", font));
            table18.addCell(getPDFCell("历史拒贷平台总数", font));
            table18.addCell(getPDFCell("今日拒贷平台个数", font));
            table18.addCell(getPDFCell("近3天拒贷平台个数", font));
            table18.addCell(getPDFCell("近7天拒贷平台个数", font));
            table18.addCell(getPDFCell("近15天拒贷平台个数", font));
            table18.addCell(getPDFCell("近30天拒贷平台个数", font));
            table18.addCell(getPDFCell("近60天拒贷平台个数", font));
            table18.addCell(getPDFCell("近90天拒贷平台个数", font));

            for(int i = 0; i < 10; i++)
            {
                table18.addCell(getPDFCell("", font));
            }

            //table19
            PdfPTable table19 = new PdfPTable(5);
            table19.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f});
            table19.setSpacingAfter(15f);
            table19.getDefaultCell().setBorder(0);
            table19.addCell(getPDFCell("序号", font));
            table19.addCell(getPDFCell("拒贷平台编号", font));
            table19.addCell(getPDFCell("平台类型", font));
            table19.addCell(getPDFCell("最早拒贷时间", font));
            table19.addCell(getPDFCell("最近拒贷时间", font));

            table19.addCell(getPDFCell("1", font));
            table19.addCell(getPDFCell("Ssw564", font));
            table19.addCell(getPDFCell("金融", font));
            table19.addCell(getPDFCell("2015/06/02 15:45:12", font));
            table19.addCell(getPDFCell("2015/06/02 15:45:12", font));

            table19.addCell(getPDFCell("2", font));
            table19.addCell(getPDFCell("SD", font));
            table19.addCell(getPDFCell("股票", font));
            table19.addCell(getPDFCell("2018/02/23 16:12:23", font));
            table19.addCell(getPDFCell("2015/06/02 15:45:12", font));

            //table20
            PdfPTable table20 = new PdfPTable(2);
            table20.setTotalWidth(new float[]{200f, 200f});
            table20.setSpacingAfter(15f);
            PdfPCell cell27 = new PdfPCell();
            cell27 = mergeCol("说明", font, 2);
            cell27.setBorder(0);
            cell27.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell28 = new PdfPCell();
            cell28 = mergeCol("1、该报告展示被查询人被其他平台拒绝放款情况;", font, 2);
            cell28.setBorder(0);
            cell28.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell29 = new PdfPCell();
            cell29 = mergeCol("2、该报告仅为自有数据统计;", font, 2);
            cell29.setBorder(0);
            cell29.setHorizontalAlignment(Element.ALIGN_LEFT);
            table20.addCell(cell27);
            table20.addCell(cell28);
            table20.addCell(cell29);

            /*
             *第6页
             *table21
             * */
            PdfPTable table21 = new PdfPTable(2);
            table21.setTotalWidth(new float[]{200f, 200f});
            table21.setSpacingAfter(10f);
            PdfPCell cell30 = new PdfPCell();
            cell30 = mergeCol("查询人:张某", font1, 2);
            cell30.setBorder(0);
            cell30.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell31 = new PdfPCell();
            cell31 = mergeCol("编号:xhsgs45", font1, 1);
            cell31.setBorder(0);
            cell31.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell32 = new PdfPCell();
            cell32 = mergeCol("2018/02/23 12:21:23", font1, 1);
            cell32.setBorder(0);
            cell32.setHorizontalAlignment(Element.ALIGN_RIGHT);

            table21.addCell(cell30);
            table21.addCell(cell31);
            table21.addCell(cell32);

            table21.addCell(getPDFCell("历史最大还款金额", font));
            table21.addCell(getPDFCell("", font));

            table21.addCell(getPDFCell("", font));
            table21.addCell(getPDFCell("", font));

            //table22
            PdfPTable table22 = new PdfPTable(10);
            table22.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f});
            table22.setSpacingAfter(10f);
            table22.addCell(getPDFCell("还款最早出现时间", font));
            table22.addCell(getPDFCell("还款最近出现时间", font));
            table22.addCell(getPDFCell("历史还款平台总数", font));
            table22.addCell(getPDFCell("今日还款平台个数", font));
            table22.addCell(getPDFCell("近3天还款平台个数", font));
            table22.addCell(getPDFCell("近7天还款平台个数", font));
            table22.addCell(getPDFCell("近15天还款平台个数", font));
            table22.addCell(getPDFCell("近30天还款平台个数", font));
            table22.addCell(getPDFCell("近60天还款平台个数", font));
            table22.addCell(getPDFCell("近90天还款平台个数", font));

            for(int i = 0; i < 10; i++)
            {
                table18.addCell(getPDFCell("", font));
            }

            //table23
            PdfPTable table23 = new PdfPTable(5);
            table23.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f});
            table23.setSpacingAfter(15f);
            table23.getDefaultCell().setBorder(0);
            table23.addCell(getPDFCell("序号", font));
            table23.addCell(getPDFCell("还款平台编号", font));
            table23.addCell(getPDFCell("平台类型", font));
            table23.addCell(getPDFCell("最早还款时间", font));
            table23.addCell(getPDFCell("最近还款时间", font));

            table23.addCell(getPDFCell("1", font));
            table23.addCell(getPDFCell("Ssw564", font));
            table23.addCell(getPDFCell("金融", font));
            table23.addCell(getPDFCell("2015/06/02 15:45:12", font));
            table23.addCell(getPDFCell("2015/06/02 15:45:12", font));

            table23.addCell(getPDFCell("2", font));
            table23.addCell(getPDFCell("SD", font));
            table23.addCell(getPDFCell("股票", font));
            table23.addCell(getPDFCell("2018/02/23 16:12:23", font));
            table23.addCell(getPDFCell("2015/06/02 15:45:12", font));

            //table24
            PdfPTable table24 = new PdfPTable(2);
            table24.setTotalWidth(new float[]{200f, 200f});
            table24.setSpacingAfter(15f);
            PdfPCell cell33 = new PdfPCell();
            cell33 = mergeCol("说明", font, 2);
            cell33.setBorder(0);
            cell33.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell34 = new PdfPCell();
            cell34 = mergeCol("1、该报告展示被查询人在其他平台还款情况;", font, 2);
            cell34.setBorder(0);
            cell34.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell35 = new PdfPCell();
            cell35 = mergeCol("2、该报告仅为自有数据统计;", font, 2);
            cell35.setBorder(0);
            cell35.setHorizontalAlignment(Element.ALIGN_LEFT);
            table24.addCell(cell33);
            table24.addCell(cell34);
            table24.addCell(cell35);

            /*
             *第7页
             *table25
             * */
            PdfPTable table25 = new PdfPTable(2);
            table25.setTotalWidth(new float[]{200f, 200f});
            table25.setSpacingAfter(10f);
            PdfPCell cell36 = new PdfPCell();
            cell36 = mergeCol("查询人:张某", font1, 2);
            cell36.setBorder(0);
            cell36.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell37 = new PdfPCell();
            cell37 = mergeCol("编号:xhsgs45", font1, 1);
            cell37.setBorder(0);
            cell37.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell38 = new PdfPCell();
            cell38 = mergeCol("2018/02/23 12:21:23", font1, 1);
            cell38.setBorder(0);
            cell38.setHorizontalAlignment(Element.ALIGN_RIGHT);

            table25.addCell(cell36);
            table25.addCell(cell37);
            table25.addCell(cell38);

            table25.addCell(getPDFCell("历史最大逾期金额", font));
            table25.addCell(getPDFCell("历史最长逾期天数", font));

            table25.addCell(getPDFCell("", font));
            table25.addCell(getPDFCell("", font));

            //table26
            PdfPTable table26 = new PdfPTable(10);
            table26.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f, 200f});
            table26.setSpacingAfter(10f);
            table26.addCell(getPDFCell("逾期最早出现时间", font));
            table26.addCell(getPDFCell("逾期最近出现时间", font));
            table26.addCell(getPDFCell("历史逾期平台总数", font));
            table26.addCell(getPDFCell("今日逾期平台个数", font));
            table26.addCell(getPDFCell("近3天逾期平台个数", font));
            table26.addCell(getPDFCell("近7天逾期平台个数", font));
            table26.addCell(getPDFCell("近15天逾期平台个数", font));
            table26.addCell(getPDFCell("近30天逾期平台个数", font));
            table26.addCell(getPDFCell("近60天逾期平台个数", font));
            table26.addCell(getPDFCell("近90天逾期平台个数", font));

            for(int i = 0; i < 10; i++)
            {
                table26.addCell(getPDFCell("", font));
            }

            //table27
            PdfPTable table27 = new PdfPTable(5);
            table27.setTotalWidth(new float[]{200f, 200f, 200f, 200f, 200f});
            table27.setSpacingAfter(15f);
            table27.getDefaultCell().setBorder(0);
            table27.addCell(getPDFCell("序号", font));
            table27.addCell(getPDFCell("逾期平台编号", font));
            table27.addCell(getPDFCell("平台类型", font));
            table27.addCell(getPDFCell("最早逾期时间", font));
            table27.addCell(getPDFCell("最近逾期时间", font));

            table27.addCell(getPDFCell("1", font));
            table27.addCell(getPDFCell("Ssw564", font));
            table27.addCell(getPDFCell("金融", font));
            table27.addCell(getPDFCell("2015/06/02 15:45:12", font));
            table27.addCell(getPDFCell("2015/06/02 15:45:12", font));

            table27.addCell(getPDFCell("2", font));
            table27.addCell(getPDFCell("SD", font));
            table27.addCell(getPDFCell("股票", font));
            table27.addCell(getPDFCell("2018/02/23 16:12:23", font));
            table27.addCell(getPDFCell("2015/06/02 15:45:12", font));

            //table28
            PdfPTable table28 = new PdfPTable(2);
            table28.setTotalWidth(new float[]{200f, 200f});
            table28.setSpacingAfter(15f);
            PdfPCell cell39 = new PdfPCell();
            cell39 = mergeCol("说明", font, 2);
            cell39.setBorder(0);
            cell39.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell40 = new PdfPCell();
            cell40 = mergeCol("1、该报告展示被查询人逾期不还的情况;", font, 2);
            cell40.setBorder(0);
            cell40.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell41 = new PdfPCell();
            cell41 = mergeCol("2、该报告仅为自有数据统计;", font, 2);
            cell41.setBorder(0);
            cell41.setHorizontalAlignment(Element.ALIGN_LEFT);
            table28.addCell(cell39);
            table28.addCell(cell40);
            table28.addCell(cell41);

            /*
             *第8页
             *table29
             * */
            PdfPTable table29 = new PdfPTable(2);
            table29.setTotalWidth(new float[]{200f, 200f});
            table29.setSpacingAfter(10f);
            PdfPCell cell42 = new PdfPCell();
            cell42 = mergeCol("查询人:张某", font1, 2);
            cell42.setBorder(0);
            cell42.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell43 = new PdfPCell();
            cell43 = mergeCol("编号:xhsgs45", font1, 1);
            cell43.setBorder(0);
            cell43.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell44 = new PdfPCell();
            cell44 = mergeCol("2018/02/23 12:21:23", font1, 1);
            cell44.setBorder(0);
            cell44.setHorizontalAlignment(Element.ALIGN_RIGHT);

            table29.addCell(cell42);
            table29.addCell(cell43);
            table29.addCell(cell44);

            //table30
            PdfPTable table30 = new PdfPTable(3);
            table30.setTotalWidth(new float[]{100f, 100f, 100f});
            table30.setSpacingAfter(50f);
            PdfPCell cell45 = new PdfPCell(new Paragraph("已被拉入黑名单", font));
            cell45.setBorder(0);
            cell45.setHorizontalAlignment(Element.ALIGN_CENTER);
            String imagePath3 = "D:\\hei.png";
            Image image3 = Image.getInstance(imagePath3);
            //image3.setWidthPercentage(100);
            image3.setAlignment(Element.ALIGN_CENTER);

            PdfPCell cellx = new PdfPCell(new Paragraph("", font));
            cellx.setBorder(0);

            table30.getDefaultCell().setBorder(0);
            table30.addCell(cellx);
            table30.addCell(image3);
            table30.addCell(cellx);

            table30.addCell(cellx);
            table30.addCell(cell45);
            table30.addCell(cellx);

            //table31
            PdfPTable table31 = new PdfPTable(3);
            table31.setTotalWidth(new float[]{100f, 100f, 100f});
            table31.setSpacingAfter(50f);
            PdfPCell cell46 = new PdfPCell(new Paragraph("非黑名单客户", font));
            cell46.setBorder(0);
            cell46.setHorizontalAlignment(Element.ALIGN_CENTER);
            String imagePath4 = "D:\\bai.png";
            Image image4 = Image.getInstance(imagePath4);
            //image4.setWidthPercentage(100);
            image4.setAlignment(Element.ALIGN_CENTER);

            table31.getDefaultCell().setBorder(0);
            table31.addCell(cellx);
            table31.addCell(image4);
            table31.addCell(cellx);

            table31.addCell(cellx);
            table31.addCell(cell46);
            table31.addCell(cellx);

            //table32
            PdfPTable table32 = new PdfPTable(2);
            table32.setTotalWidth(new float[]{200f, 200f});
            table32.setSpacingAfter(15f);
            PdfPCell cell47 = new PdfPCell();
            cell47 = mergeCol("说明", font, 2);
            cell47.setBorder(0);
            cell47.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell48 = new PdfPCell();
            cell48 = mergeCol("1、该报告展示被查询人逾期不还的情况;", font, 2);
            cell48.setBorder(0);
            cell48.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell49 = new PdfPCell();
            cell49 = mergeCol("2、该报告仅为自有数据统计;", font, 2);
            cell49.setBorder(0);
            cell49.setHorizontalAlignment(Element.ALIGN_LEFT);
            table32.addCell(cell47);
            table32.addCell(cell48);
            table32.addCell(cell49);

            /*
             *第9页
             *table33
             * */
            PdfPTable table33 = new PdfPTable(2);
            table33.setTotalWidth(new float[]{200f, 200f});
            table33.setSpacingAfter(15f);
            PdfPCell cell50 = new PdfPCell();
            cell50 = mergeCol("1、该报告产生的结果仅为对元数据自有数据源进行分析得出，未持有数据无法分析;", font, 2);
            cell50.setBorder(0);
            cell50.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell51 = new PdfPCell();
            cell51 = mergeCol("2、元数据不保证该报告的真实性和准确性，但承诺在信息汇总、加工、整合过程中保持客观、中立的地位;", font, 2);
            cell51.setBorder(0);
            cell51.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell52 = new PdfPCell();
            cell52 = mergeCol("3、该报告为机密文件，仅限于您在已授权情况下了解授权人信用状况下使用，不得泄露;", font, 2);
            cell52.setBorder(0);
            cell52.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell53 = new PdfPCell();
            cell53 = mergeCol("4、若为取得被查人授权的情况下，查询该报告或者将该报告用作他用，造成的一切后果均由违规查询和使用者承担;", font, 2);
            cell53.setBorder(0);
            cell53.setHorizontalAlignment(Element.ALIGN_LEFT);
            PdfPCell cell54 = new PdfPCell();
            cell54 = mergeCol("5、更多说明或有疑问，请联系：400 001 5351;", font, 2);
            cell54.setBorder(0);
            cell54.setHorizontalAlignment(Element.ALIGN_LEFT);
            table33.addCell(cell50);
            table33.addCell(cell51);
            table33.addCell(cell52);
            table33.addCell(cell53);
            table33.addCell(cell54);

            /*
             * 开始写入pdf
             * */
            document.open();

            //为报告添加页脚
            PdfPTable pdfPTable = new PdfPTable(1);
            Footer footerTable = new Footer(pdfPTable);
            footerTable.setTableFooter(writer, font0);
            //为报告添加页眉
            Header headerTable = new Header(pdfPTable);
            headerTable.setTableHeader(writer, font0);
            document.add(pdfPTable);

            Paragraph title1 = new Paragraph("个人信用报告", font2);
            title1.setSpacingBefore(50f);
            title1.setAlignment(Paragraph.ALIGN_CENTER);
            Paragraph title2 = new Paragraph("(明细)", font);
            title2.setSpacingBefore(10f);
            title2.setAlignment(Paragraph.ALIGN_CENTER);

            table2.setTableEvent(event);
            table3.setTableEvent(event);
            table4.setTableEvent(event);
            table5.setTableEvent(event);
            table6.setTableEvent(event);
            table7.setTableEvent(event);

            document.add(table1);
            document.add(title1);
            document.add(title2);
            document.add(table2);
            document.add(table3);
            document.add(table4);
            document.add(table5);
            document.add(table6);
            document.add(table7);
            document.add(table8);

            /*
             *第3页
             * */
            document.newPage();
            //为报告添加页脚
            PdfPTable pdfPTable1 = new PdfPTable(1);
            Footer footerTable1 = new Footer(pdfPTable1);
            footerTable1.setTableFooter(writer, font0);
            //为报告添加页眉
            Header headerTable1 = new Header(pdfPTable1);
            headerTable1.setTableHeader(writer, font0);
            document.add(pdfPTable1);

            Paragraph title3 = new Paragraph("--天镜系列·行为报告--", font3);
            title3.setSpacingAfter(25f);
            title3.setAlignment(Paragraph.ALIGN_CENTER);
            Paragraph title4 = new Paragraph("天镜·多头客", font2);
            title4.setSpacingAfter(20f);
            title4.setAlignment(Paragraph.ALIGN_CENTER);
            Paragraph title5 = new Paragraph("注册平台展示", font1);
            title5.setSpacingAfter(15f);
            title5.setAlignment(Paragraph.ALIGN_CENTER);

            table10.setTableEvent(event);
            table11.setTableEvent(event);

            document.add(title3);
            document.add(title4);
            document.add(table9);
            document.add(table10);
            document.add(title5);
            document.add(table11);
            document.add(table12);

            /*
             *第4页
             * */
            document.newPage();
            //为报告添加页脚
            PdfPTable pdfPTable2 = new PdfPTable(1);
            Footer footerTable2 = new Footer(pdfPTable2);
            footerTable2.setTableFooter(writer, font0);
            //为报告添加页眉
            Header headerTable2 = new Header(pdfPTable2);
            headerTable2.setTableHeader(writer, font0);
            document.add(pdfPTable2);

            Paragraph title6 = new Paragraph("天镜·通过客", font2);
            title6.setSpacingAfter(20f);
            title6.setAlignment(Paragraph.ALIGN_CENTER);
            Paragraph title7 = new Paragraph("通过贷款平台展示", font1);
            title7.setSpacingAfter(15f);
            title7.setAlignment(Paragraph.ALIGN_CENTER);

            table14.setTableEvent(event);
            table15.setTableEvent(event);

            document.add(title6);
            document.add(table13);
            document.add(table14);
            document.add(title7);
            document.add(table15);
            document.add(table16);

            /*
             *第5页
             * */
            document.newPage();
            //为报告添加页脚
            PdfPTable pdfPTable3 = new PdfPTable(1);
            Footer footerTable3 = new Footer(pdfPTable3);
            footerTable3.setTableFooter(writer, font0);
            //为报告添加页眉
            Header headerTable3 = new Header(pdfPTable3);
            headerTable3.setTableHeader(writer, font0);
            document.add(pdfPTable3);

            Paragraph title8 = new Paragraph("天镜·拒贷客", font2);
            title8.setSpacingAfter(20f);
            title8.setAlignment(Paragraph.ALIGN_CENTER);
            Paragraph title9 = new Paragraph("拒贷平台展示", font1);
            title9.setSpacingAfter(15f);
            title9.setAlignment(Paragraph.ALIGN_CENTER);

            table18.setTableEvent(event);
            table19.setTableEvent(event);

            document.add(title8);
            document.add(table17);
            document.add(table18);
            document.add(title9);
            document.add(table19);
            document.add(table20);

            /*
             *第6页
             * */
            document.newPage();
            //为报告添加页脚
            PdfPTable pdfPTable4 = new PdfPTable(1);
            Footer footerTable4 = new Footer(pdfPTable4);
            footerTable4.setTableFooter(writer, font0);
            //为报告添加页眉
            Header headerTable4 = new Header(pdfPTable4);
            headerTable4.setTableHeader(writer, font0);
            document.add(pdfPTable4);

            Paragraph title10 = new Paragraph("天镜·优良客", font2);
            title10.setSpacingAfter(20f);
            title10.setAlignment(Paragraph.ALIGN_CENTER);
            Paragraph title11 = new Paragraph("还款平台展示", font1);
            title11.setSpacingAfter(15f);
            title11.setAlignment(Paragraph.ALIGN_CENTER);

            table22.setTableEvent(event);
            table23.setTableEvent(event);

            document.add(title10);
            document.add(table21);
            document.add(table22);
            document.add(title11);
            document.add(table23);
            document.add(table24);

            /*
             *第7页
             * */
            document.newPage();
            //为报告添加页脚
            PdfPTable pdfPTable5 = new PdfPTable(1);
            Footer footerTable5 = new Footer(pdfPTable5);
            footerTable5.setTableFooter(writer, font0);
            //为报告添加页眉
            Header headerTable5 = new Header(pdfPTable5);
            headerTable5.setTableHeader(writer, font0);
            document.add(pdfPTable5);

            Paragraph title12 = new Paragraph("天镜·常欠客", font2);
            title12.setSpacingAfter(20f);
            title12.setAlignment(Paragraph.ALIGN_CENTER);
            Paragraph title13 = new Paragraph("逾期平台展示", font1);
            title13.setSpacingAfter(15f);
            title13.setAlignment(Paragraph.ALIGN_CENTER);

            table26.setTableEvent(event);
            table27.setTableEvent(event);

            document.add(title12);
            document.add(table25);
            document.add(table26);
            document.add(title13);
            document.add(table27);
            document.add(table28);

            /*
             *第8页
             * */
            document.newPage();
            //为报告添加页脚
            PdfPTable pdfPTable6 = new PdfPTable(1);
            Footer footerTable6 = new Footer(pdfPTable6);
            footerTable6.setTableFooter(writer, font0);
            //为报告添加页眉
            Header headerTable6 = new Header(pdfPTable6);
            headerTable6.setTableHeader(writer, font0);
            document.add(pdfPTable6);

            Paragraph title14 = new Paragraph("天镜·黑名单", font2);
            title14.setSpacingAfter(20f);
            title14.setAlignment(Paragraph.ALIGN_CENTER);
            document.add(title14);
            document.add(table29);
            document.add(table30);
            document.add(table31);
            document.add(table32);

            /*
             *第9页
             * */
            document.newPage();
            //为报告添加页脚
            PdfPTable pdfPTable7 = new PdfPTable(1);
            Footer footerTable7 = new Footer(pdfPTable7);
            footerTable7.setTableFooter(writer, font0);
            //为报告添加页眉
            Header headerTable7 = new Header(pdfPTable7);
            headerTable7.setTableHeader(writer, font0);
            document.add(pdfPTable7);

            Paragraph title15 = new Paragraph("报告说明", font2);
            title15.setSpacingAfter(20f);
            title15.setAlignment(Paragraph.ALIGN_CENTER);

            document.add(title15);
            document.add(table33);
            document.close();
            zip.closeEntry();
        }
        zip.close();
    }

    //合并行的静态函数
    public static PdfPCell mergeRow(String str, Font font, int i)
    {
        //创建单元格对象，将内容及字体传入
        PdfPCell cell = new PdfPCell(new Paragraph(str, font));

        //设置单元格内容居中
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);

        //将该单元格所在列包括单元格在内的i行单元格合并为一个单元格
        cell.setRowspan(i);
        return cell;
    }

    //和并列的静态函数
    public static PdfPCell mergeCol(String str, Font font, int i)
    {
        PdfPCell cell = new PdfPCell(new Paragraph(str, font));
        cell.setMinimumHeight(25);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);

        //将该单元格所在列包括单元格在内的i行单元格合并为一个单元格
        cell.setColspan(i);
        return cell;
    }

    //获取指定内容与字体的单元格
    public static PdfPCell getPDFCell(String string, Font font)
    {
        PdfPCell cell = new PdfPCell(new Paragraph(string, font));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);

        //设置最小单元格高度
        cell.setMinimumHeight(25);
        return cell;
    }

    //页脚事件
    private static class Footer extends PdfPageEventHelper
    {
        public static PdfPTable footer;

        @SuppressWarnings("static-access")
        public Footer(PdfPTable footer)
        {
            this.footer = footer;
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document)
        {
            //把页脚表格定位
            footer.writeSelectedRows(0, -1, 38, 50, writer.getDirectContent());
        }

        /*页脚是文字*/
        public void setTableFooter(PdfWriter writer, Font font)
        {
            PdfPTable table = new PdfPTable(1);
            table.setTotalWidth(520f);
            PdfPCell cell = new PdfPCell();
            cell.setBorder(1);
            String string = "本报告为机密文件,请勿泄露";
            Paragraph p = new Paragraph(string, font);
            p.setAlignment(Paragraph.ALIGN_CENTER);
            cell.setPaddingLeft(10f);
            cell.setPaddingTop(-2f);
            cell.addElement(p);
            table.addCell(cell);
            Footer event = new Footer(table);
            writer.setPageEvent(event);
        }
    }

    //页眉事件
    private static class Header extends PdfPageEventHelper
    {
        public static PdfPTable header;

        public Header(PdfPTable header)
        {
            Header.header = header;
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document)
        {
            //把页眉表格定位
            header.writeSelectedRows(0, -1, 60, 806, writer.getDirectContent());
        }

        /*设置页眉*/
        public void setTableHeader(PdfWriter writer, Font font)
                throws MalformedURLException, IOException, DocumentException
        {
            String imageAddress = "D:\\";
            PdfPTable table = new PdfPTable(2);
            table.setTotalWidth(new float[]{200f, 283f});
            /*table.setTotalWidth(555);*/
            PdfPCell cell = new PdfPCell();
            PdfPCell cell2 = new PdfPCell();
            Paragraph p = new Paragraph("专业、高效、安全的金融风控管理平台", font);
            p.setAlignment(Element.ALIGN_RIGHT);
            cell2.setBorder(0);
            cell.setBorder(0);
            Image image01;
            image01 = Image.getInstance(imageAddress + "yuan.png"); //图片自己传
            image01.setAlignment(Paragraph.ALIGN_LEFT);
            //image01.scaleAbsolute(355f, 10f);
            image01.setWidthPercentage(30);
            cell.setPaddingLeft(10f);
            cell.setPaddingTop(-20f);
            cell.addElement(image01);
            cell2.setPaddingRight(10f);
            cell2.setPaddingTop(-20f);
            cell2.addElement(p);
            table.addCell(cell);
            table.addCell(cell2);
            Header event = new Header(table);
            writer.setPageEvent(event);
        }
    }

    //页码事件
    private static class PageXofYTest extends PdfPageEventHelper
    {

        public PdfTemplate total;

        public BaseFont bfChinese;

        /*重写PdfPageEventHelper中的onOpenDocument方法*/
        @Override
        public void onOpenDocument(PdfWriter writer, Document document)
        {
            // 得到文档的内容并为该内容新建一个模板
            total = writer.getDirectContent().createTemplate(500, 500);
            try
            {

                String prefixFont = "";
                String os = System.getProperties().getProperty("os.name");
                if(os.startsWith("win") || os.startsWith("Win"))
                {
                    prefixFont = "C:\\Windows\\Fonts" + File.separator;
                }
                else
                {
                    prefixFont = "/usr/share/fonts/chinese" + File.separator;
                }

                // 设置字体对象为Windows系统默认的字体
                bfChinese = BaseFont.createFont(prefixFont + "simsun.ttc,0", BaseFont.IDENTITY_H,
                        BaseFont.NOT_EMBEDDED);
            }
            catch(Exception e)
            {
                throw new ExceptionConverter(e);
            }
        }

        /*重写PdfPageEventHelper中的onEndPage方法*/
        @Override
        public void onEndPage(PdfWriter writer, Document document)
        {
            // 新建获得用户页面文本和图片内容位置的对象
            PdfContentByte pdfContentByte = writer.getDirectContent();
            // 保存图形状态
            pdfContentByte.saveState();
            String text = writer.getPageNumber() + "/";
            // 获取点字符串的宽度
            float textSize = bfChinese.getWidthPoint(text, 9);
            pdfContentByte.beginText();
            // 设置随后的文本内容写作的字体和字号
            pdfContentByte.setFontAndSize(bfChinese, 9);

            // 定位'X/'
            float x = (document.right() + document.left()) / 2;
            float y = 56f;
            pdfContentByte.setTextMatrix(x, y);
            pdfContentByte.showText(text);
            pdfContentByte.endText();

            // 将模板加入到内容（content）中- // 定位'Y'
            pdfContentByte.addTemplate(total, x + textSize, y);

            pdfContentByte.restoreState();
        }

        /*重写PdfPageEventHelper中的onCloseDocument方法*/
        @Override
        public void onCloseDocument(PdfWriter writer, Document document)
        {
            total.beginText();
            try
            {
                String prefixFont = "";
                String os = System.getProperties().getProperty("os.name");
                if(os.startsWith("win") || os.startsWith("Win"))
                {
                    prefixFont = "C:\\Windows\\Fonts" + File.separator;
                }
                else
                {
                    prefixFont = "/usr/share/fonts/chinese" + File.separator;
                }

                bfChinese = BaseFont.createFont(prefixFont + "simsun.ttc,0", BaseFont.IDENTITY_H,
                        BaseFont.NOT_EMBEDDED);
                total.setFontAndSize(bfChinese, 9);
            }
            catch(DocumentException e)
            {
                e.printStackTrace();
            }
            catch(IOException e)
            {
                e.printStackTrace();
            }
            total.setTextMatrix(0, 0);
            // 设置总页数的值到模板上，并应用到每个界面
            total.showText(String.valueOf(writer.getPageNumber() - 1));
            total.endText();
        }
    }
}
