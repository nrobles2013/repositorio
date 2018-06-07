/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.procesos;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Image;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.GrayColor;
import java.io.ByteArrayOutputStream;
import java.nio.file.Paths;
import java.util.List;
import javax.ejb.Stateless;
import pop.comun.dominio.MaeTipoCambio;
import pop.webcobranzas.util.HeaderFooter;

/**
 *
 * @author Jyoverar
 */
@Stateless(mappedName = "ejbTipoCambioRep")
public class SessionTipoCambioRep implements IRepTipoCambio {

    public static final Font FONT = new Font(FontFamily.HELVETICA, 12, Font.NORMAL, GrayColor.GRAYWHITE);

    @Override
    public byte[] imprimirTipoCambio(List<MaeTipoCambio> oMaeTipoCambios) throws Exception {
        System.out.println("pop.webcobranzas.procesos.SessionTipoCambioRep.imprimirTipoCambio()");

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        Document document = new Document(PageSize.A4);
        PdfWriter writer = PdfWriter.getInstance(document, baos);

        writer.setPageEvent(new HeaderFooter());

        document.open();
        PdfContentByte cb = writer.getDirectContentUnder();

        String imagenPath = Paths.get("/pop/webcobranzas/resources/jpg", "Koala.jpg").toString();
        System.out.println(imagenPath);

        //Image jpg = null;
        //System.out.println(this.getClass().getClassLoader().getResource(imagenPath));
        //jpg = Image.getInstance(imagenPath);
        //document.add(getWatermarkedImage(cb, jpg,""));
        document.add(new Paragraph(" Hello Jhon Yovera ramos ... por finnnnn"));
        document.newPage();
        
        float fntSize, lineSpacing;
        fntSize = 6.7f;
        lineSpacing = 10f;
        
        Paragraph p = new Paragraph(new Phrase(lineSpacing, "Hello Jhon Yovera ramos ... por finnnnn",
                FontFactory.getFont(FontFactory.COURIER, fntSize)));
        
        document.add(p);
        
        document.newPage();
        document.add(new Paragraph("Hello World222222!"));

        document.close();
        return baos.toByteArray();

    }

    public Image getWatermarkedImage(PdfContentByte cb, Image img, String watermark) throws DocumentException {
        float width = 50;//img.getScaledWidth();
        float height = 20;//img.getScaledHeight();
        PdfTemplate template = cb.createTemplate(width, height);
        template.addImage(img, width, 0, 0, height, 0, 0);
        ColumnText.showTextAligned(template, Element.ALIGN_CENTER,
                new Phrase(watermark, FONT), width / 2, height / 2, 30);
        return Image.getInstance(template);
    }

}
