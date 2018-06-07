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
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;
import java.util.Locale;
import java.util.UUID;
import javax.ejb.Stateless;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import pop.comun.dominio.MaeCronograma;
import pop.comun.dominio.MaeInversion;
import pop.comun.dominio.MaeReporte;
import pop.comun.dominio.reporte.RepSaldoDeudor;
import pop.webcobranzas.util.HeaderFooter;

/**
 *
 * @author Jyoverar
 */
@Stateless(mappedName = "ejbSaldoDeudorRep")
public class SessionSaldoDeudor implements IRepSaldoDeudor {

    @Override
    public byte[] imprimirSaldoDeudor(RepSaldoDeudor oSaldoDeudor, MaeReporte oMaeReporte) throws Exception {

        //System.out.println("pop.webcobranzas.procesos.SessionSaldoDeudor.imprimirSaldoDeudor() i");
        // para los numeros 
        Locale.setDefault(new Locale("en", "US"));
        // configurando la pagina
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        Document document = new Document(PageSize.A4, 36, 36, 170, 36);
        PdfWriter writer = PdfWriter.getInstance(document, baos);
        // obteniendo las fuentes
        String fontCalibriPath = Paths.get("/pop/webcobranzas/resources/font", "calibri.ttf").toString();
        String fontCalibriPathB = Paths.get("/pop/webcobranzas/resources/font", "calibrib.ttf").toString();
        BaseFont bf = BaseFont.createFont(fontCalibriPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        BaseFont bfb = BaseFont.createFont(fontCalibriPathB, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font font;
        Font fontB;
        BaseColor bColor = new BaseColor(28, 50, 93);
        BaseColor bColorBorde;
        // para los formatos de los numero y fechas
        DecimalFormat formatterNum = new DecimalFormat("###,###,##0.00");
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        // celda universal
        PdfPCell cell;
        // primera tabla de cabecera ( 3 columnas)
        PdfPTable table = new PdfPTable(3);
        table.setTotalWidth(500);
        table.setWidths(new float[]{50, 10, 40});

        PdfPTable tableCab = new PdfPTable(1);
        // codigo
        fontB = new Font(bfb, 12);
        cell = new PdfPCell(new Paragraph(12, "CÓDIGO: " + oSaldoDeudor.getMaeInversion().getCInversion().trim(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        // dni
        fontB = new Font(bfb, 10);
        cell = new PdfPCell(new Paragraph(10, "DNI: " + oSaldoDeudor.getMaeInversion().getcPersonaId().getANroDocumento(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        // datos de la persona
        cell = new PdfPCell(new Paragraph(10, oSaldoDeudor.getMaeInversion().getcPersonaId().getDApePat() + " "
                + oSaldoDeudor.getMaeInversion().getcPersonaId().getDApeMat() + " "
                + oSaldoDeudor.getMaeInversion().getcPersonaId().getDNombres(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        font = new Font(bf, 8);
        // direccion
        cell = new PdfPCell(new Paragraph(12, "PREDIO: " + oSaldoDeudor.getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getADir1()
                + "  " + oSaldoDeudor.getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoProv().trim()
                + " - " + oSaldoDeudor.getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist().trim(), font));
        cell.setBorder(0);
        tableCab.addCell(cell);
        // quitando los bordes
        PdfPCell wrongCell = new PdfPCell(tableCab);
        wrongCell.setBorder(Rectangle.NO_BORDER);
        // adicionando la primera tabla convertida en wrong a la tabla Columna 1
        table.addCell(wrongCell);
        // adicionando la primera tabla convertida en wrong a la tabla Columna 2 en blanco
        wrongCell = new PdfPCell(new Phrase(""));
        wrongCell.setBorder(Rectangle.NO_BORDER);
        table.addCell(wrongCell);
        // creando la tabla de la columna 3 con 2 columnas 
        PdfPTable tableB = new PdfPTable(2);
        tableB.setWidths(new float[]{70, 30});
        //right.setIndentationLeft(10);
        font = new Font(bf, 10);
        bColorBorde = new BaseColor(28, 50, 93);

        cell = new PdfPCell(new Paragraph(12, "Fecha inico de la operación:", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatter.format(oSaldoDeudor.getMaeInversion().getFEmision()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Fecha fin de la operación:", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatter.format(oSaldoDeudor.getMaeInversion().getFVencimiento()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Tasa de la operación:", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, Double.toString((Double) oSaldoDeudor.getMaeInversion().getPTasa() * 100) + "%", font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Monto de la operación:", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format((Double) oSaldoDeudor.getMaeInversion().getIInversion()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBorderColor(bColorBorde);
        cell.setBorderWidth(new Float(1.5));
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableB.addCell(cell);

        wrongCell = new PdfPCell(tableB);
        wrongCell.setBorder(Rectangle.NO_BORDER);
        table.addCell(wrongCell);

        // creando la cabecera y el pie de pagina
        HeaderFooter headerFooter = new HeaderFooter(table);
        if (oSaldoDeudor.getMaeInversion().getMaeFondo().getCFondoId().equals("0001")) {
            //headerFooter.setNameLogo("logoemprendedor.png");
            oMaeReporte.setNameLogo("logoemprendedor.png");
        } else if (oSaldoDeudor.getMaeInversion().getMaeFondo().getCFondoId().equals("0002")) {
            //headerFooter.setNameLogo("logopopular.png");
            oMaeReporte.setNameLogo("logopopular.png");
        } else if (oSaldoDeudor.getMaeInversion().getMaeFondo().getCFondoId().equals("0003")) {
            //headerFooter.setNameLogo("logomype.png");
            oMaeReporte.setNameLogo("logomype.png");
        } else {
            //headerFooter.setNameLogo("logosafi.png");
            oMaeReporte.setNameLogo("logosafi.png");
        }
        // datos del reporte
        //oMaeReporte.setNameReport("  ESTADO DE CUENTA A LA FECHA " + formatter.format(oMaeReporte.getfIniBusq()));
        oMaeReporte.setNameReport("SALDO DEUDOR DEL " + formatter.format(oSaldoDeudor.getFactual()) + "  AL " + formatter.format(oSaldoDeudor.getFfutura()));
        //headerFooter.setNameReport("SALDO DEUDOR DEL " + formatter.format(oSaldoDeudor.getFactual()) + "  AL " + formatter.format(oSaldoDeudor.getFfutura()));
        //headerFooter.setFecha(formatter.format(oMaeReporte.getfIniBusq()));
        //headerFooter.setUserName(oMaeReporte.getcUsuarioAdd());
        headerFooter.setMaeReporte(oMaeReporte);

        // aumentando la cabecera y detalle
        writer.setPageEvent(headerFooter);
        document.open();
        //document.add(table);
        //font = new Font(bf, 9, Font.NORMAL, BaseColor.WHITE);

        document.add(new Paragraph(" ", font));

        PdfPTable tableDetalle = new PdfPTable(3);
        tableDetalle.setWidthPercentage(90);
        tableDetalle.setWidths(new float[]{30, 40, 30});

        font = new Font(bf, 9, Font.NORMAL, BaseColor.WHITE);
        bColor = new BaseColor(28, 50, 93);
        bColorBorde = new BaseColor(28, 50, 93);
        cell = new PdfPCell(new Paragraph(11, "Fecha de emisión del reporte", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(11, "Fecha de cumplimiento de compromiso", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(11, "Fecha de la última cuota generada", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);

        font = new Font(bf, 9, Font.NORMAL, BaseColor.BLACK);
        bColor = new BaseColor(255, 255, 255);
        bColorBorde = new BaseColor(200, 200, 200);
        cell = new PdfPCell(new Paragraph(11, formatter.format(oSaldoDeudor.getFactual()), font));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        //
        cell = new PdfPCell(new Paragraph(11, formatter.format(oSaldoDeudor.getFfutura()), font));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        //
        if (oSaldoDeudor.getFultCuota() != null) {
            cell = new PdfPCell(new Paragraph(11, formatter.format(oSaldoDeudor.getFultCuota()), font));
        }else{
            cell = new PdfPCell(new Paragraph(11, "No Generada", font));
        }
        
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_CENTER);
        tableDetalle.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);

        document.add(tableDetalle);
        document.add(new Paragraph(" ", font));

        //------------------------------------------------- detalle
        tableDetalle = new PdfPTable(2);
        tableDetalle.setWidthPercentage(90);
        tableDetalle.setWidths(new float[]{30, 70});
        // a favor del fondo
        font = new Font(bf, 10, Font.NORMAL, BaseColor.WHITE);
        bColor = new BaseColor(28, 50, 93);
        bColorBorde = new BaseColor(28, 50, 93);
        cell = new PdfPCell(new Paragraph(12, "A favor del fondo:", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        tableDetalle.addCell(cell);
        //
        font = new Font(bf, 10, Font.NORMAL, BaseColor.BLACK);
        bColor = new BaseColor(255, 255, 255);
        bColorBorde = new BaseColor(200, 200, 200);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getNfavorFondo()), font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        tableDetalle.addCell(cell);

        document.add(tableDetalle);
        document.add(new Paragraph(" ", font));

        // -- tabla 
        tableDetalle = new PdfPTable(3);
        tableDetalle.setWidthPercentage(90);
        tableDetalle.setWidths(new float[]{60, 2, 38});
        // tabla de cuotas vencidas
        PdfPTable tableCv = new PdfPTable(5);
        tableCv.setWidthPercentage(100);
        tableCv.setWidths(new float[]{18, 18, 28, 18, 18});

        font = new Font(bf, 9, Font.NORMAL, BaseColor.BLACK);
        bColor = new BaseColor(182, 221, 240);
        bColorBorde = new BaseColor(200, 200, 200);
        cell = new PdfPCell(new Paragraph(12, "Cuotas vencidas adeudadas", font));
        cell.setColspan(5);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableCv.addCell(cell);

        font = new Font(bf, 9, Font.NORMAL, BaseColor.BLACK);

        cell = new PdfPCell(new Paragraph(12, "N° Cuota", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Capital", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Interes Compensatorio", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Interes moratorio", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Total", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableCv.addCell(cell);
        tableCv.setHeaderRows(1);

        bColor = new BaseColor(255, 255, 255);
        bColorBorde = new BaseColor(200, 200, 200);

        String flColor = "0";
        for (MaeCronograma mc : oSaldoDeudor.getMaeCronogramaList()) {
            if (flColor.equals("0")) {
                bColor = new BaseColor(240, 240, 240);
            } else {
                bColor = new BaseColor(255, 255, 255);
            }

            cell = new PdfPCell(new Paragraph(12, "" + mc.getnSecuencia(), font));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableCv.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, formatterNum.format(mc.getIcapital()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableCv.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, formatterNum.format(mc.getIinteres()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableCv.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, formatterNum.format(mc.getImora()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableCv.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, formatterNum.format(((Double) mc.getIcapital()) + ((Double) mc.getIinteres()) + ((Double) mc.getImora())), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableCv.addCell(cell);
            if (flColor.equals("0")) {
                flColor = "1";
            } else {
                flColor = "0";
            }
        }
        // totales
        cell = new PdfPCell(new Paragraph(12, "Total cuotas vencidas", font));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIcapAtra()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIicAtra()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIimAtra()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getNtotAtra()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);
        // primera columna
        wrongCell = new PdfPCell(tableCv);
        wrongCell.setBorder(Rectangle.NO_BORDER);
        tableDetalle.addCell(wrongCell);
        // en blanco -- 2 columna
        wrongCell = new PdfPCell(new Phrase(""));
        wrongCell.setBorder(Rectangle.NO_BORDER);
        tableDetalle.addCell(wrongCell);

        // cuotas por vencer
        tableCv = new PdfPTable(2);
        tableCv.setWidthPercentage(100);
        tableCv.setWidths(new float[]{70, 30});
        font = new Font(bf, 9, Font.NORMAL, BaseColor.BLACK);
        bColor = new BaseColor(182, 221, 240);
        bColorBorde = new BaseColor(200, 200, 200);

        cell = new PdfPCell(new Paragraph(12, "Cuotas por vencer ", font));
        cell.setColspan(2);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableCv.addCell(cell);
        //
        bColor = new BaseColor(255, 255, 255);
        bColorBorde = new BaseColor(200, 200, 200);
        cell = new PdfPCell(new Paragraph(12, "Capital ", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIcapFut()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Interes compensatorio del " + formatter.format(oSaldoDeudor.getFultCuota()) + " al " + formatter.format(oSaldoDeudor.getFfutura()), font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIicFut()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);

        cell = new PdfPCell(new Paragraph(12, "Total por vencer ", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(((Double) oSaldoDeudor.getIcapFut()) + ((Double) oSaldoDeudor.getIicFut())), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableCv.addCell(cell);
        // ultima columna 3
        wrongCell = new PdfPCell(tableCv);
        wrongCell.setBorder(Rectangle.NO_BORDER);
        tableDetalle.addCell(wrongCell);
        document.add(tableDetalle);
        document.add(new Paragraph(" ", font));
        // sobrecargo bancario
        tableDetalle = new PdfPTable(2);
        tableDetalle.setWidthPercentage(80);
        tableDetalle.setWidths(new float[]{50, 50});

        cell = new PdfPCell(new Paragraph(12, "Sobrecargo bancario", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIcargCuoAtra()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Gastos legales", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIgastLegalFut()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Gastos administrativos", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIgastAdmin()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableDetalle.addCell(cell);

        document.add(tableDetalle);
        document.add(new Paragraph(" ", font));

        tableDetalle = new PdfPTable(2);
        tableDetalle.setWidthPercentage(90);
        tableDetalle.setWidths(new float[]{30, 70});
        font = new Font(bf, 10, Font.NORMAL, BaseColor.WHITE);
        bColor = new BaseColor(0, 102, 51);
        bColorBorde = new BaseColor(0, 102, 51);
        cell = new PdfPCell(new Paragraph(12, "A favor del cliente:", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        tableDetalle.addCell(cell);
        font = new Font(bf, 10, Font.NORMAL, BaseColor.BLACK);
        bColor = new BaseColor(255, 255, 255);
        bColorBorde = new BaseColor(200, 200, 200);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getIsaldoFavor()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableDetalle.addCell(cell);
        document.add(tableDetalle);
        document.add(new Paragraph(" ", font));
        //  total deuda
        tableDetalle = new PdfPTable(2);
        tableDetalle.setWidthPercentage(90);
        tableDetalle.setWidths(new float[]{30, 70});
        font = new Font(bf, 10, Font.NORMAL, BaseColor.WHITE);
        bColor = new BaseColor(153, 0, 0);
        bColorBorde = new BaseColor(153, 0, 0);
        cell = new PdfPCell(new Paragraph(12, "Total de la deuda:", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        tableDetalle.addCell(cell);
        font = new Font(bf, 10, Font.NORMAL, BaseColor.BLACK);
        bColor = new BaseColor(255, 255, 255);
        bColorBorde = new BaseColor(200, 200, 200);
        cell = new PdfPCell(new Paragraph(12, formatterNum.format(oSaldoDeudor.getNtotDebe() + oSaldoDeudor.getIgastLegalFut() + oSaldoDeudor.getIgastAdmin()), font));
        cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        tableDetalle.addCell(cell);
        document.add(tableDetalle);
        document.add(new Paragraph(" ", font));
        
        document.addAuthor(oMaeReporte.getUserName());
        document.addCreationDate();
        document.addCreator("popular-safi.com");
        document.addTitle("SALDO DEUDOR DEL " + formatter.format(oSaldoDeudor.getFactual()) + "  AL " + formatter.format(oSaldoDeudor.getFfutura()));
        
        // grabando archivo para autidoria
        Date date = new Date();
        LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        String ruta;
        String uuid;
        uuid = oSaldoDeudor.getMaeInversion().getCInversion().trim() + "_" + UUID.randomUUID().toString() + ".pdf";
        ruta = "D:\\webcobranzas\\files\\debit_balance\\" + localDate.getYear()
                + "\\" + localDate.getMonthValue()
                + "\\" + localDate.getDayOfMonth()
                + "\\" + uuid;
        File fileO = new File(ruta);
        fileO.getParentFile().mkdirs();
        FileOutputStream fos = new FileOutputStream(fileO);
        
        document.close();
        baos.writeTo(fos);
        fos.close();
        //System.out.println("pop.webcobranzas.procesos.SessionSaldoDeudor.imprimirSaldoDeudor() f");
        return baos.toByteArray();

    }

    @Override
    public byte[] exportarSaldoDeudor(RepSaldoDeudor oSaldoDeudor, MaeReporte oMaeReporte) throws Exception {
        //System.out.println(" <i> pop.webcobranzas.procesos.SessionSaldoDeudor.exportarSaldoDeudor() ");
        // logo
        String nameLogo = "";
        String nameFondo = "";

        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
        int xLine = 3;

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        // font negrilla
        org.apache.poi.ss.usermodel.Font fontNeg = workbook.createFont();
        fontNeg.setBoldweight(org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD);
        byte[] rgb = new byte[3];
        rgb[0] = (byte) 166; // red
        rgb[1] = (byte) 112; // green
        rgb[2] = (byte) 12; // blue

        XSSFColor myColor = new XSSFColor(rgb);

        // style black
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(fontNeg);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

        CellStyle headerStyleRigth = workbook.createCellStyle();
        headerStyleRigth.setAlignment(CellStyle.ALIGN_RIGHT);
        headerStyleRigth.setFont(fontNeg);
        headerStyleRigth.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyleRigth.setFillPattern(CellStyle.SOLID_FOREGROUND);
        headerStyleRigth.setDataFormat((short) 7);

        CellStyle bodyStyle = workbook.createCellStyle();

        CellStyle bodyStyleRigth = workbook.createCellStyle();
        bodyStyleRigth.setAlignment(CellStyle.ALIGN_RIGHT);

        DataFormat datafrmt = workbook.createDataFormat();
        //bodyStyleRigth.setDataFormat(datafrmt.getFormat("_ S/. * #,##0_ ;_ S/. * -#,##0_ ;_ S/. * \"-\"_ ;_ @_ "));
        bodyStyleRigth.setDataFormat((short) 7);

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");

        if (oSaldoDeudor == null) {
            return byteOutputStream.toByteArray();
        }

        switch (oSaldoDeudor.getMaeInversion().getMaeFondo().getCFondoId()) {
            case "0001":
                nameLogo = "logoemprendedor.png";
                nameFondo = "Fondo Capital Emprendedor";
                break;
            case "0002":
                nameLogo = "logopopular.png";
                nameFondo = "Fondo Popular";
                break;
            case "0003":
                nameLogo = "logomype.png";
                nameFondo = "Fondo MYPE";
                break;
            case "0004":
                nameLogo = "logosafi.png";
                nameFondo = "Fondo Perez Hidalgo";
                break;
            default:
                break;
        }

        String imagenPath = Paths.get("/pop/webcobranzas/resources/jpg", nameLogo).toString();
        InputStream inputStream = new FileInputStream(imagenPath);
        byte[] bytes = IOUtils.toByteArray(inputStream);
        int pictureIdx = workbook.addPicture(bytes, workbook.PICTURE_TYPE_PNG);
        inputStream.close();

        CreationHelper helper = workbook.getCreationHelper();

        Drawing drawing = sheet.createDrawingPatriarch();

        ClientAnchor anchor = helper.createClientAnchor();

        anchor.setCol1(1);
        anchor.setRow1(0);
        anchor.setCol2(2); //Column C
        anchor.setRow2(1); //Row 4

        //Creates a picture
        Picture pict = drawing.createPicture(anchor, pictureIdx);
        //Reset the image to the original size
        pict.resize();
        pict.resize(0.3);

        // obteniendo la inversion
        MaeInversion inversion = oSaldoDeudor.getMaeInversion();
        // nombre de la hoja del excel
        workbook.setSheetName(0, inversion.getCInversion().trim());

        XSSFCell cell;

        XSSFRow rowHeader;

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Código");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(3);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(inversion.getCInversion().trim());
        cell = rowHeader.createCell(4);
        cell.setCellStyle(headerStyle);

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Cliente");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(3);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(inversion.getcPersonaId().getDApePat() + " "
                + inversion.getcPersonaId().getDApeMat() + " "
                + inversion.getcPersonaId().getDNombres());
        cell = rowHeader.createCell(4);
        cell.setCellStyle(headerStyle);

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Fondo");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(3);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(nameFondo.toUpperCase());
        cell = rowHeader.createCell(4);
        cell.setCellStyle(headerStyle);

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Fecha de solicitud");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(3);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFactual()));
        cell = rowHeader.createCell(4);
        cell.setCellStyle(headerStyle);

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Fecha del Estado deudor");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(3);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
        cell = rowHeader.createCell(4);
        cell.setCellStyle(headerStyle);

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue("Saldo de Capital  de Cronograma al");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
        cell = rowHeader.createCell(4);
        cell.setCellStyle(bodyStyleRigth);
        cell.setCellValue(oSaldoDeudor.getIcapFut());

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue("Compensatorio de la cuota  al");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
        cell = rowHeader.createCell(4);
        cell.setCellStyle(bodyStyleRigth);
        cell.setCellValue(oSaldoDeudor.getIicFut());

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Total saldo Cronograma");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(3);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(4);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(((Double) oSaldoDeudor.getIcapFut()) + ((Double) oSaldoDeudor.getIicFut()));

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue("Saldo de capital atrasado al");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
        cell = rowHeader.createCell(4);
        cell.setCellStyle(bodyStyleRigth);
        cell.setCellValue((double) oSaldoDeudor.getIcapAtra());

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue("Saldos de interés compensatorio atrasado al");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
        cell = rowHeader.createCell(4);
        cell.setCellStyle(bodyStyleRigth);
        cell.setCellValue((double) oSaldoDeudor.getIicAtra());

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue("Saldo a favor en EECC");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
        cell = rowHeader.createCell(4);
        cell.setCellStyle(bodyStyleRigth);
        cell.setCellValue((double) oSaldoDeudor.getIsaldoFavor());

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue("Saldos de interés moratorio");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
        cell = rowHeader.createCell(4);
        cell.setCellStyle(bodyStyleRigth);
        cell.setCellValue((double) oSaldoDeudor.getIimAtra());

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Total saldo Estado de Cuenta");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(3);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(4);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(((Double) oSaldoDeudor.getIcapAtra())
                + ((Double) oSaldoDeudor.getIicAtra())
                + ((Double) oSaldoDeudor.getIimAtra())
                - ((Double) oSaldoDeudor.getIsaldoFavor()));

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(bodyStyle);
        cell.setCellValue("Gastos de cobranza (c.notariales, protesto, otros)");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(bodyStyle);
        cell = rowHeader.createCell(4);
        cell.setCellStyle(bodyStyleRigth);
        cell.setCellValue((double) ((Double) oSaldoDeudor.getIcargCuoAtra())
                + ((Double) oSaldoDeudor.getIgastLegalFut())
                + ((Double) oSaldoDeudor.getIgastAdmin())
        );
        xLine++;

        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Deuda Total");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
        cell = rowHeader.createCell(3);
        cell.setCellStyle(headerStyle);
        cell = rowHeader.createCell(4);
        cell.setCellStyle(headerStyleRigth);
        cell.setCellValue(oSaldoDeudor.getNtotDebe() + oSaldoDeudor.getIgastLegalFut() + oSaldoDeudor.getIgastAdmin());

        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);

        workbook.write(byteOutputStream);
        
        // grabando archivo para autidoria
        Date date = new Date();
        LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        String ruta;
        String uuid;
        uuid = inversion.getCInversion().trim() + "_" + UUID.randomUUID().toString() + ".xlsx";
        ruta = "D:\\webcobranzas\\files\\debit_balance\\" + localDate.getYear()
                + "\\" + localDate.getMonthValue()
                + "\\" + localDate.getDayOfMonth()
                + "\\" + uuid;
        File fileO = new File(ruta);
        fileO.getParentFile().mkdirs();
        FileOutputStream fos = new FileOutputStream(fileO);
        byteOutputStream.writeTo(fos);
        fos.close();
        
        //System.out.println(" <f> pop.webcobranzas.procesos.SessionSaldoDeudor.exportarSaldoDeudor() ");
        return byteOutputStream.toByteArray();
    }

//    public byte[] exportarSaldoDeudorNo(RepSaldoDeudor oSaldoDeudor, MaeReporte oMaeReporte) throws Exception {
//        System.out.println(" <i> pop.webcobranzas.procesos.SessionSaldoDeudor.exportarSaldoDeudor() ");
//
//        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
//        int xLine = 1;
//
//        XSSFWorkbook workbook = new XSSFWorkbook();
//        XSSFSheet sheet = workbook.createSheet();
//        // font negrilla
//        org.apache.poi.ss.usermodel.Font fontNeg = workbook.createFont();
//        fontNeg.setBoldweight(org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD);
//        // style black
//        CellStyle headerStyle = workbook.createCellStyle();
//        headerStyle.setFont(fontNeg);
//        // style
//        CellStyle style = workbook.createCellStyle();
//
//        //DecimalFormat formatterNum = new DecimalFormat("###,###,###.00");
//        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
//
//        if (oSaldoDeudor == null) {
//            return byteOutputStream.toByteArray();
//        }
//        // obteniendo la inversion
//        MaeInversion inversion = oSaldoDeudor.getMaeInversion();
//        // nombre de la hoja del excel
//        workbook.setSheetName(0, inversion.getCInversion().trim());
//
//        XSSFCell cell;
//
//        // codigo
//        XSSFRow rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("CÓDIGO:");
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue(inversion.getCInversion().trim());
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("Fecha inico de la operación:");
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue(formatter.format(inversion.getFEmision()));
//        // dni
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("DNI:");
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue(inversion.getcPersonaId().getANroDocumento());
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("Fecha fin de la operación:");
//        rowHeader.createCell(9).setCellValue(formatter.format(inversion.getFVencimiento()));
//        // persona
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("APELLIDOS Y NOMBRES:");
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue(inversion.getcPersonaId().getDApePat() + " "
//                + inversion.getcPersonaId().getDApeMat() + " "
//                + inversion.getcPersonaId().getDNombres());
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("Tasa de la operación:");
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue((double) inversion.getPTasa() * 100 + "%");
//        // direccion
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("PREDIO:");
//        if (inversion.getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist() != null) {
//            cell = rowHeader.createCell(2);
//            cell.setCellStyle(headerStyle);
//            cell.setCellValue(inversion.getcPersonaId().getMaeDireccionList().get(0).getADir1()
//                    + "  " + inversion.getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoProv().trim()
//                    + " - " + inversion.getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist().trim());
//        } else {
//            rowHeader.createCell(1).setCellValue(" ");
//        }
//        
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue("Monto de la operación:");
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(headerStyle);
//        cell.setCellValue((double) inversion.getIInversion());
//
//        sheet.addMergedRegion(new CellRangeAddress(4, 4, 2, 7));
//        
//        xLine++;
//
//        // style black
//        CellStyle headerStyleCenter = workbook.createCellStyle();
//        headerStyleCenter.setAlignment(CellStyle.ALIGN_CENTER);
//        headerStyleCenter.setBorderBottom(CellStyle.BORDER_THIN);
//        headerStyleCenter.setBorderLeft(CellStyle.BORDER_THIN);
//        headerStyleCenter.setBorderRight(CellStyle.BORDER_THIN);
//        headerStyleCenter.setBorderTop(CellStyle.BORDER_THIN);
//        headerStyleCenter.setFont(fontNeg);
//
//        // fechas de informacion
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue("Fecha de emisión del reporte");
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue("Fecha de cumplimiento de compromiso");
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(6);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue("Fecha de la última cuota generada");
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(headerStyleCenter);
//
//        rowHeader = sheet.createRow((short) xLine++);
//
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue(formatter.format(oSaldoDeudor.getFactual()));
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue(formatter.format(oSaldoDeudor.getFfutura()));
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(6);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue(formatter.format(oSaldoDeudor.getFultCuota()));
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(headerStyleCenter);
//
//        sheet.addMergedRegion(new CellRangeAddress(6, 6, 1, 3));
//        sheet.addMergedRegion(new CellRangeAddress(6, 6, 4, 6));
//        sheet.addMergedRegion(new CellRangeAddress(6, 6, 7, 9));
//
//        sheet.addMergedRegion(new CellRangeAddress(7, 7, 1, 3));
//        sheet.addMergedRegion(new CellRangeAddress(7, 7, 4, 6));
//        sheet.addMergedRegion(new CellRangeAddress(7, 7, 7, 9));
//
//        xLine++;
//        // style black
//        CellStyle headerStyleRigth = workbook.createCellStyle();
//        headerStyleRigth.setAlignment(CellStyle.ALIGN_RIGHT);
//        headerStyleRigth.setBorderBottom(CellStyle.BORDER_THIN);
//        headerStyleRigth.setBorderLeft(CellStyle.BORDER_THIN);
//        headerStyleRigth.setBorderRight(CellStyle.BORDER_THIN);
//        headerStyleRigth.setBorderTop(CellStyle.BORDER_THIN);
//        headerStyleRigth.setFont(fontNeg);
//
//        CellStyle bodyStyle = workbook.createCellStyle();
//        bodyStyle.setBorderBottom(CellStyle.BORDER_THIN);
//        bodyStyle.setBorderLeft(CellStyle.BORDER_THIN);
//        bodyStyle.setBorderRight(CellStyle.BORDER_THIN);
//        bodyStyle.setBorderTop(CellStyle.BORDER_THIN);
//
//        CellStyle bodyStyleCenter = workbook.createCellStyle();
//        bodyStyleCenter.setAlignment(CellStyle.ALIGN_CENTER);
//        bodyStyleCenter.setBorderBottom(CellStyle.BORDER_THIN);
//        bodyStyleCenter.setBorderLeft(CellStyle.BORDER_THIN);
//        bodyStyleCenter.setBorderRight(CellStyle.BORDER_THIN);
//        bodyStyleCenter.setBorderTop(CellStyle.BORDER_THIN);
//
//        CellStyle bodyStyleRigth = workbook.createCellStyle();
//        bodyStyleRigth.setAlignment(CellStyle.ALIGN_RIGHT);
//        bodyStyleRigth.setBorderBottom(CellStyle.BORDER_THIN);
//        bodyStyleRigth.setBorderLeft(CellStyle.BORDER_THIN);
//        bodyStyleRigth.setBorderRight(CellStyle.BORDER_THIN);
//        bodyStyleRigth.setBorderTop(CellStyle.BORDER_THIN);
//
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue("A favor del fondo");
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(headerStyleRigth);
//        cell.setCellValue(oSaldoDeudor.getNfavorFondo());
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(headerStyleRigth);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(headerStyleRigth);
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(headerStyleRigth);
//        cell = rowHeader.createCell(6);
//        cell.setCellStyle(headerStyleRigth);
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(headerStyleRigth);
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(headerStyleRigth);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(headerStyleRigth);
//        sheet.addMergedRegion(new CellRangeAddress(9, 9, 1, 8));
//
//        xLine++;
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue("Cuotas vencidas adeudadas");
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(headerStyleCenter);
//
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(headerStyleCenter);
//        cell.setCellValue("Cuotas por vencer");
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(headerStyleCenter);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(headerStyleCenter);
//
//        sheet.addMergedRegion(new CellRangeAddress(11, 11, 1, 5));
//        sheet.addMergedRegion(new CellRangeAddress(11, 11, 7, 9));
//
//        String[] headers = new String[]{
//            "N° Cuota", "Capital", "Interes Compensatorio", "Interes moratorio", "Total"
//        };
//
//        rowHeader = sheet.createRow((short) xLine++);
//        for (int i = 0; i < headers.length; ++i) {
//            String header = headers[i];
//            cell = rowHeader.createCell(i + 1);
//            cell.setCellStyle(headerStyleCenter);
//            cell.setCellValue(header);
//        }
//
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(bodyStyle);
//        cell.setCellValue("Capital");
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue(oSaldoDeudor.getIcapFut());
//
//        int xLineB = xLine - 1;
//        System.out.println("pop.webcobranzas.procesos.SessionSaldoDeudor.exportarSaldoDeudor()---" + xLineB);
//
//        int flg = 0;
//
//        for (MaeCronograma mc : oSaldoDeudor.getMaeCronogramaList()) {
//            rowHeader = sheet.createRow((short) xLine++);
//            // N° Cuota
//            cell = rowHeader.createCell(1);
//            cell.setCellStyle(bodyStyleCenter);
//            cell.setCellValue((int) mc.getnSecuencia());
//            // Capital
//            cell = rowHeader.createCell(2);
//            cell.setCellStyle(bodyStyleRigth);
//            cell.setCellValue((double) mc.getIcapital());
//            // Interes Compensatorio
//            cell = rowHeader.createCell(3);
//            cell.setCellStyle(bodyStyleRigth);
//            cell.setCellValue((double) mc.getIinteres());
//            // Interes moratorio
//            cell = rowHeader.createCell(4);
//            cell.setCellStyle(bodyStyleRigth);
//            cell.setCellValue((double) mc.getImora());
//            // Total
//            cell = rowHeader.createCell(5);
//            cell.setCellStyle(bodyStyleRigth);
//            cell.setCellValue((double) ((Double) mc.getIcapital()) + ((Double) mc.getIinteres()) + ((Double) mc.getImora()));
//
//            if (flg == 0) {
//                cell = rowHeader.createCell(7);
//                cell.setCellStyle(bodyStyle);
//                cell.setCellValue("Interes compensatorio del " + formatter.format(oSaldoDeudor.getFultCuota()) + " al " + formatter.format(oSaldoDeudor.getFfutura()));
//                cell = rowHeader.createCell(8);
//                cell.setCellStyle(bodyStyle);
//                cell = rowHeader.createCell(9);
//                cell.setCellStyle(bodyStyleRigth);
//                cell.setCellValue(oSaldoDeudor.getIicFut());
//            }
//
//            if (flg == 1) {
//                cell = rowHeader.createCell(7);
//                cell.setCellStyle(bodyStyle);
//                cell.setCellValue("Total por vencer");
//                cell = rowHeader.createCell(8);
//                cell.setCellStyle(bodyStyle);
//                cell = rowHeader.createCell(9);
//                cell.setCellStyle(bodyStyleRigth);
//                cell.setCellValue(((Double) oSaldoDeudor.getIcapFut()) + ((Double) oSaldoDeudor.getIicFut()));
//            }
//            flg++;
//        }
//
//        sheet.addMergedRegion(new CellRangeAddress(12, 12, 7, 8));
//        sheet.addMergedRegion(new CellRangeAddress(13, 13, 7, 8));
//        sheet.addMergedRegion(new CellRangeAddress(14, 14, 7, 8));
//
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(bodyStyleCenter);
//        cell.setCellValue("Total cuotas vencidas");
//
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double) oSaldoDeudor.getIcapAtra());
//
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double) oSaldoDeudor.getIicAtra());
//
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double) oSaldoDeudor.getIimAtra());
//
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double) oSaldoDeudor.getNtotAtra());
//
//        xLine++;
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(bodyStyle);
//        cell.setCellValue("Sobracargo bancario");
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(6);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double) oSaldoDeudor.getIcargCuoAtra());
//        
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(bodyStyle);
//        cell.setCellValue("Gatos legales");
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(6);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double) oSaldoDeudor.getIgastLegalFut());
//        
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(bodyStyle);
//        cell.setCellValue("Gastos administrativos");
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(6);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double) oSaldoDeudor.getIgastAdmin());
//
//        sheet.addMergedRegion(new CellRangeAddress(18, 18, 2, 8));
//        sheet.addMergedRegion(new CellRangeAddress(19, 19, 2, 8));
//        sheet.addMergedRegion(new CellRangeAddress(20, 20, 2, 8));
//        
//        xLine++;
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(bodyStyle);
//        cell.setCellValue("A favor del cliente");
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(6);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double)oSaldoDeudor.getIsaldoFavor());
//        
//        xLine++;
//        rowHeader = sheet.createRow((short) xLine++);
//        cell = rowHeader.createCell(1);
//        cell.setCellStyle(bodyStyle);
//        cell.setCellValue("Total de la deuda");
//        cell = rowHeader.createCell(2);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(3);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(4);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(5);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(6);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(7);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(8);
//        cell.setCellStyle(bodyStyle);
//        cell = rowHeader.createCell(9);
//        cell.setCellStyle(bodyStyleRigth);
//        cell.setCellValue((double)oSaldoDeudor.getNtotDebe() + oSaldoDeudor.getIgastLegalFut() + oSaldoDeudor.getIgastAdmin());
//        
//        sheet.addMergedRegion(new CellRangeAddress(22, 22, 1, 8));
//        sheet.addMergedRegion(new CellRangeAddress(24, 24, 1, 8));
//        
//        workbook.write(byteOutputStream);
//        System.out.println(" <f> pop.webcobranzas.procesos.SessionSaldoDeudor.exportarSaldoDeudor() ");
//        return byteOutputStream.toByteArray();
//    }
//
//    
}
