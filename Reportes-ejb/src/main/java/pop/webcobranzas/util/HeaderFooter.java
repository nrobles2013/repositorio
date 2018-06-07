/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.util;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.ExceptionConverter;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.GrayColor;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.ElementList;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import java.io.IOException;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.logging.Level;
import java.util.logging.Logger;
import pop.comun.dominio.MaeReporte;

/**
 *
 * @author Jyoverar
 */
public class HeaderFooter extends PdfPageEventHelper {

    protected ElementList header;
    protected ElementList footer;
    protected PdfPTable headerTable;

    public static final String HEADER = "<table width=\"100%\" border=\"0\"><tr>"
            + "<td>Header </td>"
            + "<td align=\"right\">Some title</td></tr></table>";
    public static final String FOOTER = "<table width=\"100%\" border=\"0\"><tr><td>Footer</td><td align=\"right\">Some title</td></tr></table>";

    public static final Font FONT = new Font(Font.FontFamily.TIMES_ROMAN, 10, Font.NORMAL, GrayColor.GRAYWHITE);
    public static final String fontCalibriPath = Paths.get("/pop/webcobranzas/resources/font", "calibri.ttf").toString();
    public static final String fontCalibriPathB = Paths.get("/pop/webcobranzas/resources/font", "calibrib.ttf").toString();

    Image jpg = null;

    private MaeReporte maeReporte;
    private String fecha;

    public HeaderFooter() throws IOException {
        header = XMLWorkerHelper.parseToElementList(HEADER, null);
        footer = XMLWorkerHelper.parseToElementList(FOOTER, null);
    }

    public HeaderFooter(PdfPTable headerTable) {
        this.headerTable = headerTable;
    }

    @Override
    public void onEndPage(PdfWriter writer, Document document) {
        try {
            // font
            BaseFont bf = BaseFont.createFont(fontCalibriPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            BaseFont bfb = BaseFont.createFont(fontCalibriPathB, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font font = new Font(bf, 10);
            Font fontB = new Font(bfb, 12);

            // cabecera azul
            ColumnText ct = new ColumnText(writer.getDirectContent());
            PdfContentByte canvas = writer.getDirectContent();
            BaseColor bColor = new BaseColor(28, 50, 93);
            Rectangle rect1 = new Rectangle(0, 810, 700, 850);
            rect1.setBackgroundColor(bColor);
            rect1.setBorder(Rectangle.BOX);
            rect1.setBorderWidth(0);
            ct.setSimpleColumn(rect1);
            canvas.rectangle(rect1);
            ct.go();
            // logo
            String imagenPath = Paths.get("/pop/webcobranzas/resources/jpg", getMaeReporte().getNameLogo()).toString();
            jpg = Image.getInstance(imagenPath);
            Element eImg = jpg;
            jpg.scaleAbsoluteHeight(60);
            jpg.scaleAbsoluteWidth(150);
            ct.setSimpleColumn(new Rectangle(36, 740, 200, 810));
            ct.addElement(eImg);
            ct.go();
            // text www 
            Phrase headerText = new Phrase("| www.popular-safi.com | Teléfono: " + getMaeReporte().getNumAsesor() + " ", FONT);
            ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT,
                    headerText,
                    400-(getMaeReporte().getNumAsesor().length()),
                    820, 0);
            ct.go();
            // text estado de cuenta
            Rectangle rect3 = new Rectangle(210, 785, 700, 810);
            rect3.setBackgroundColor(new BaseColor(209, 211, 212));
            rect3.setBorder(Rectangle.BOX);
            rect3.setBorderWidth(0);
            ct.setSimpleColumn(rect3);
            canvas.rectangle(rect3);
            ct.go();
            // text estado de cuenta
            ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT,
                    new Phrase( getMaeReporte().getNameReport(), fontB),
                    300, 790, 0);
            ct.go();
            // text pagina
            ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT,
                    new Phrase("Pag.: " + writer.getCurrentPageNumber(), font),
                    500, 770, 0);
            ct.go();
            // tabla       
            headerTable.writeSelectedRows(0, -1, 60, 750, writer.getDirectContent());
            // pie azul
            Rectangle rect2 = new Rectangle(0, 0, 700, 40);
            rect2.setBackgroundColor(bColor);
            rect2.setBorder(Rectangle.BOX);
            rect2.setBorderWidth(0);
            ct.setSimpleColumn(rect2);
            canvas.rectangle(rect2);
            ct.go();
            // texto pie
            font = new Font(bf, 8, Font.NORMAL, BaseColor.WHITE);
            LocalDateTime today = LocalDateTime.now();
            ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT,
                    new Phrase( ""+today, font),
                    36, 10, 0);
            ct.go();
            ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT,
                    new Phrase( getMaeReporte().getMailAsesor(), font),
                    36, 20, 0);
            ct.go();
            ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT,
                    new Phrase( getMaeReporte().getUserName(), font),
                    36, 30, 0);
            ct.go();
//            ct.setSimpleColumn(new Rectangle(36, 10, 559, 32));
////            for (Element e : footer) {
////                ct.addElement(e);
////            }
//            ct.go();

        } catch (DocumentException de) {
            throw new ExceptionConverter(de);
        } catch (IOException ex) {
            Logger.getLogger(HeaderFooter.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public String getFecha() {
        return fecha;
    }

    public void setFecha(String fecha) {
        this.fecha = fecha;
    }

    public MaeReporte getMaeReporte() {
        return maeReporte;
    }

    public void setMaeReporte(MaeReporte maeReporte) {
        this.maeReporte = maeReporte;
    }

}
