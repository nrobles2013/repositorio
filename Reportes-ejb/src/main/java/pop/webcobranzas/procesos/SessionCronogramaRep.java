/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.procesos;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.UUID;
import javax.ejb.Stateless;
import pop.comun.dominio.MaeCronograma;
import pop.comun.dominio.MaeReporte;
import pop.webcobranzas.util.HeaderFooter;

/**
 *
 * @author Jyoverar
 */
@Stateless(mappedName = "ejbCronogramaRep")
public class SessionCronogramaRep implements IRepCronograma {

    @Override
    public byte[] imprimirCronograma(List<MaeCronograma> oMaeCronogramas, MaeReporte maeReporte) throws Exception {
        //System.out.println("<i> Reporete de crongrama");

        Locale.setDefault(new Locale("en", "US"));

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        Document document = new Document(PageSize.A4, 36, 36, 170, 36);
        PdfWriter writer = PdfWriter.getInstance(document, baos);

        String fontCalibriPath = Paths.get("/pop/webcobranzas/resources/font", "calibri.ttf").toString();
        String fontCalibriPathB = Paths.get("/pop/webcobranzas/resources/font", "calibrib.ttf").toString();
        BaseFont bf = BaseFont.createFont(fontCalibriPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        BaseFont bfb = BaseFont.createFont(fontCalibriPathB, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font font;
        Font fontB;
        BaseColor bColor = new BaseColor(28, 50, 93);
        BaseColor bColorBorde;

        DecimalFormat formatterNum = new DecimalFormat("###,###,###.00");
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");

        PdfPCell cell;

        PdfPTable table = new PdfPTable(3);

        PdfPTable tableCab = new PdfPTable(1);

        table.setTotalWidth(500);
        //table.setWidthPercentage(80);
        table.setWidths(new float[]{50, 10, 40});
        fontB = new Font(bfb, 12);
        cell = new PdfPCell(new Paragraph(12, "CÓDIGO: " + oMaeCronogramas.get(0).getMaeInversion().getCInversion().trim(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        //
        fontB = new Font(bfb, 10);
        cell = new PdfPCell(new Paragraph(10, "DNI: " + oMaeCronogramas.get(0).getMaeInversion().getcPersonaId().getANroDocumento(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        //
        cell = new PdfPCell(new Paragraph(10, oMaeCronogramas.get(0).getMaeInversion().getcPersonaId().getDApePat() + " "
                + oMaeCronogramas.get(0).getMaeInversion().getcPersonaId().getDApeMat() + " "
                + oMaeCronogramas.get(0).getMaeInversion().getcPersonaId().getDNombres(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        font = new Font(bf, 8);

        if (oMaeCronogramas.get(0).getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist() != null) {
            cell = new PdfPCell(new Paragraph(12, oMaeCronogramas.get(0).getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getADir1()
                    + "  " + oMaeCronogramas.get(0).getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoProv().trim()
                    + " - " + oMaeCronogramas.get(0).getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist().trim(), font));
        } else {
            cell = new PdfPCell(new Paragraph(12, " ", font));
        }

        cell.setBorder(0);
        tableCab.addCell(cell);
        //wrong.setIndentationLeft(10);
        PdfPCell wrongCell = new PdfPCell(tableCab);
        wrongCell.setBorder(Rectangle.NO_BORDER);

        table.addCell(wrongCell);
        wrongCell = new PdfPCell(new Phrase(""));
        wrongCell.setBorder(Rectangle.NO_BORDER);
        table.addCell(wrongCell);

        //
        PdfPTable tableB = new PdfPTable(2);
        //right.setIndentationLeft(10);
        font = new Font(bf, 10);
        bColorBorde = new BaseColor(28, 50, 93);

        cell = new PdfPCell(new Paragraph(12, "Fecha Inicio:", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatter.format(oMaeCronogramas.get(0).getMaeInversion().getFEmision()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Fecha Fin:", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatter.format(oMaeCronogramas.get(0).getMaeInversion().getFVencimiento()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Tasa:", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, Double.toString((Double) oMaeCronogramas.get(0).getMaeInversion().getPTasa() * 100) + "%", font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Monto:", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format((Double) oMaeCronogramas.get(0).getMaeInversion().getIInversion()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);

        wrongCell = new PdfPCell(tableB);
        wrongCell.setBorder(Rectangle.NO_BORDER);
        table.addCell(wrongCell);

        table.addCell(wrongCell);

        HeaderFooter headerFooter = new HeaderFooter(table);

        if (oMaeCronogramas.get(0).getMaeInversion().getMaeFondo().getCFondoId().equals("0001")) {
            //headerFooter.setNameLogo("logoemprendedor.png");
            maeReporte.setNameLogo("logoemprendedor.png");
        } else if (oMaeCronogramas.get(0).getMaeInversion().getMaeFondo().getCFondoId().equals("0002")) {
            //headerFooter.setNameLogo("logopopular.png");
            maeReporte.setNameLogo("logopopular.png");
        } else if (oMaeCronogramas.get(0).getMaeInversion().getMaeFondo().getCFondoId().equals("0003")) {
            //headerFooter.setNameLogo("logomype.png");
            maeReporte.setNameLogo("logomype.png");
        } else {
            //headerFooter.setNameLogo("logosafi.png");
            maeReporte.setNameLogo("logosafi.png");
        }
        maeReporte.setNameReport(" CRONOGRAMA DE PAGO ");
        //headerFooter.setNameReport("  ESTADO DE CUENTA A LA FECHA " + formatter.format(maeReporte.getfIniBusq()));
        //headerFooter.setFecha(formatter.format(maeReporte.getfIniBusq()));
        //headerFooter.setUserName(maeReporte.getcUsuarioAdd());
        headerFooter.setMaeReporte(maeReporte);

        writer.setPageEvent(headerFooter);
        document.open();

        font = new Font(bf, 9, Font.NORMAL, BaseColor.WHITE);

        //document.add(new Paragraph("  a", font));
        PdfPTable tableDetalle = new PdfPTable(6);

        tableDetalle.setWidthPercentage(80);
        tableDetalle.setWidths(new float[]{10, 15, 15, 20, 20, 20});

        bColorBorde = new BaseColor(200, 200, 200);

        cell = new PdfPCell(new Paragraph(12, "Cuota", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Fecha pago", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Capital", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Interes", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Cuota a Depositar", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Saldo Adeudado", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);

        tableDetalle.setHeaderRows(1);

        font = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);

        String flColor = "0";

        for (MaeCronograma crono : oMaeCronogramas) {
            if (flColor.equals("0")) {
                bColor = new BaseColor(240, 240, 240);
            } else {
                bColor = new BaseColor(255, 255, 255);
            }

            cell = new PdfPCell(new Paragraph(12, "" + crono.getnSecuencia(), font));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, formatter.format(crono.getFpago()), font));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, "" + formatterNum.format(crono.getIcapital()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, "" + formatterNum.format(crono.getIinteres()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, "" + formatterNum.format(crono.getIdeposito()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, "" + formatterNum.format(crono.getIsaldo()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);
            if (flColor.equals("0")) {
                flColor = "1";
            } else {
                flColor = "0";
            }
        }
        document.add(tableDetalle);

        document.addAuthor(maeReporte.getUserName());
        document.addCreationDate();
        document.addCreator("popular-safi.com");
        document.addTitle("CRONOGRAMA DE PAGO " + oMaeCronogramas.get(0).getMaeInversion().getCInversion().trim());
        //document.addSubject("An example to show how attributes can be added to pdf files.");

        // grabando archivo para autidoria
        Date date = new Date();
        LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        String ruta;
        String uuid;
        uuid = oMaeCronogramas.get(0).getMaeInversion().getCInversion().trim() + "_" + UUID.randomUUID().toString() + ".pdf";
        ruta = "D:\\webcobranzas\\files\\schedule\\" + localDate.getYear()
                + "\\" + localDate.getMonthValue()
                + "\\" + localDate.getDayOfMonth()
                + "\\" + uuid;
        File fileO = new File(ruta);
        fileO.getParentFile().mkdirs();
        FileOutputStream fos = new FileOutputStream(fileO);

        document.close();
        baos.writeTo(fos);
        fos.close();

        //System.out.println("<f> Reporete de crongrama");
        return baos.toByteArray();
    }

    // Add business logic below. (Right-click in editor and choose
    // "Insert Code > Add Business Method")
}
