/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.procesos;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.Font;

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
import org.apache.poi.POIXMLProperties;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import pop.comun.dominio.MaeEstadoCuenta;
import pop.comun.dominio.MaeInversion;
import pop.comun.dominio.MaeReporte;
import pop.webcobranzas.util.HeaderFooter;

/**
 *
 * @author Jyoverar
 */
@Stateless(mappedName = "ejbEstadoCuentaRep")
public class SessionEstadoCuentaRep implements IRepEstadoCuenta {

    @Override
    public byte[] imprimirEstadoCuenta(List<MaeEstadoCuenta> oMaeEstadoCuentas, MaeReporte maeReporte) throws Exception {
        //System.out.println("pop.webcobranzas.procesos.SessionEstadoCuentaRep.imprimirEstadoCuenta() i ");

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
        cell = new PdfPCell(new Paragraph(12, "CÓDIGO: " + oMaeEstadoCuentas.get(0).getMaeInversion().getCInversion().trim(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        //
        fontB = new Font(bfb, 10);
        cell = new PdfPCell(new Paragraph(10, "DNI: " + oMaeEstadoCuentas.get(0).getMaeInversion().getcPersonaId().getANroDocumento(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        //
        cell = new PdfPCell(new Paragraph(10, oMaeEstadoCuentas.get(0).getMaeInversion().getcPersonaId().getDApePat() + " "
                + oMaeEstadoCuentas.get(0).getMaeInversion().getcPersonaId().getDApeMat() + " "
                + oMaeEstadoCuentas.get(0).getMaeInversion().getcPersonaId().getDNombres(), fontB));
        cell.setBorder(0);
        tableCab.addCell(cell);
        font = new Font(bf, 8);

        if (oMaeEstadoCuentas.get(0).getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist() != null) {
            cell = new PdfPCell(new Paragraph(12, oMaeEstadoCuentas.get(0).getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getADir1()
                    + "  " + oMaeEstadoCuentas.get(0).getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoProv().trim()
                    + " - " + oMaeEstadoCuentas.get(0).getMaeInversion().getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist().trim(), font));
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
        cell = new PdfPCell(new Paragraph(12, formatter.format(oMaeEstadoCuentas.get(0).getMaeInversion().getFEmision()), font));
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
        cell = new PdfPCell(new Paragraph(12, formatter.format(oMaeEstadoCuentas.get(0).getMaeInversion().getFVencimiento()), font));
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
        cell = new PdfPCell(new Paragraph(12, Double.toString((Double) oMaeEstadoCuentas.get(0).getMaeInversion().getPTasa() * 100) + "%", font));
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
        cell = new PdfPCell(new Paragraph(12, formatterNum.format((Double) oMaeEstadoCuentas.get(0).getMaeInversion().getIInversion()), font));
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

        if (oMaeEstadoCuentas.get(0).getMaeInversion().getMaeFondo().getCFondoId().equals("0001")) {
            //headerFooter.setNameLogo("logoemprendedor.png");
            maeReporte.setNameLogo("logoemprendedor.png");
        } else if (oMaeEstadoCuentas.get(0).getMaeInversion().getMaeFondo().getCFondoId().equals("0002")) {
            //headerFooter.setNameLogo("logopopular.png");
            maeReporte.setNameLogo("logopopular.png");
        } else if (oMaeEstadoCuentas.get(0).getMaeInversion().getMaeFondo().getCFondoId().equals("0003")) {
            //headerFooter.setNameLogo("logomype.png");
            maeReporte.setNameLogo("logomype.png");
        } else {
            //headerFooter.setNameLogo("logosafi.png");
            maeReporte.setNameLogo("logosafi.png");
        }
        maeReporte.setNameReport("  ESTADO DE CUENTA A LA FECHA " + formatter.format(maeReporte.getfIniBusq()));
        //headerFooter.setNameReport("  ESTADO DE CUENTA A LA FECHA " + formatter.format(maeReporte.getfIniBusq()));
        //headerFooter.setFecha(formatter.format(maeReporte.getfIniBusq()));
        //headerFooter.setUserName(maeReporte.getcUsuarioAdd());
        headerFooter.setMaeReporte(maeReporte);

        writer.setPageEvent(headerFooter);
        document.open();

        //document.add(table);
        font = new Font(bf, 9, Font.NORMAL, BaseColor.WHITE);

        //document.add(new Paragraph("  a", font));
        PdfPTable tableDetalle = new PdfPTable(7);

        tableDetalle.setWidthPercentage(100);
        tableDetalle.setWidths(new float[]{5, 7, 13, 39, 12, 12, 12});

        bColorBorde = new BaseColor(200, 200, 200);

        cell = new PdfPCell(new Paragraph(12, "N°", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Cuota", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Fec. Operación", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Desc. Operación", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Abono", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Cargo", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        cell = new PdfPCell(new Paragraph(12, "Saldo", font));
        cell.setBackgroundColor(bColor);
        cell.setBorderColor(bColorBorde);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        tableDetalle.addCell(cell);
        tableDetalle.setHeaderRows(1);

        font = new Font(bf, 8, Font.NORMAL, BaseColor.BLACK);

        String flColor = "0";

        for (MaeEstadoCuenta estCta : oMaeEstadoCuentas) {
            if (flColor.equals("0")) {
                bColor = new BaseColor(240, 240, 240);
            } else {
                bColor = new BaseColor(255, 255, 255);
            }

            cell = new PdfPCell(new Paragraph(12, "" + estCta.getNsecuencia(), font));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);
            if (estCta.getDcuota() == null) {
                cell = new PdfPCell(new Paragraph(12, "", font));
            } else {
                cell = new PdfPCell(new Paragraph(12, "" + estCta.getDcuota(), font));
            }
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);

            cell = new PdfPCell(new Paragraph(12, formatter.format(estCta.getFoperacion()), font));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);
            cell = new PdfPCell(new Paragraph(12, estCta.getDdescriocion(), font));
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);
            cell = new PdfPCell(new Paragraph(12, formatterNum.format(estCta.getIabono()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);
            cell = new PdfPCell(new Paragraph(12, formatterNum.format(estCta.getIcargo()), font));
            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cell.setBackgroundColor(bColor);
            cell.setBorderColor(bColorBorde);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            tableDetalle.addCell(cell);
            cell = new PdfPCell(new Paragraph(12, formatterNum.format(estCta.getIsaldo()), font));
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
        document.addTitle("ESTADO DE CUENTA A LA FECHA " + formatter.format(maeReporte.getfIniBusq()) + " " + oMaeEstadoCuentas.get(0).getMaeInversion().getCInversion().trim());
        //document.addSubject("An example to show how attributes can be added to pdf files.");

        // grabando archivo para autidoria
        Date date = new Date();
        LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        String ruta;
        String uuid;
        uuid = oMaeEstadoCuentas.get(0).getMaeInversion().getCInversion().trim() + "_" + UUID.randomUUID().toString() + ".pdf";
        ruta = "D:\\webcobranzas\\files\\account_status\\" + localDate.getYear()
                + "\\" + localDate.getMonthValue()
                + "\\" + localDate.getDayOfMonth()
                + "\\" + uuid;
        File fileO = new File(ruta);
        fileO.getParentFile().mkdirs();
        FileOutputStream fos = new FileOutputStream(fileO);

//        document.newPage();
//
//        float fntSize, lineSpacing;
//        fntSize = 6.7f;
//        lineSpacing = 10f;
//
//        Paragraph p = new Paragraph(new Phrase(lineSpacing, "Hello Jhon Yovera ramos ... por finnnnn",
//                FontFactory.getFont(FontFactory.COURIER, fntSize)));
//
//        document.add(p);
        document.close();
        baos.writeTo(fos);
        fos.close();

        //System.out.println("pop.webcobranzas.procesos.SessionEstadoCuentaRep.imprimirEstadoCuenta() f ");
        return baos.toByteArray();
    }

    // Add business logic below. (Right-click in editor and choose
    // "Insert Code > Add Business Method")
    @Override
    public byte[] exportartEstadoCuenta(List<MaeEstadoCuenta> oMaeEstadoCuentas, MaeReporte maeReporte) throws Exception {

        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
        int xLine = 1;

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        // font negrilla
        org.apache.poi.ss.usermodel.Font fontNeg = workbook.createFont();
        fontNeg.setBoldweight(org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD);
        // style black
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(fontNeg);
        // style
        CellStyle style = workbook.createCellStyle();
        //style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        //style.setFillPattern(CellStyle.SOLID_FOREGROUND);

        //DecimalFormat formatterNum = new DecimalFormat("###,###,###.00");
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");

        if (oMaeEstadoCuentas == null || oMaeEstadoCuentas.isEmpty()) {
            return byteOutputStream.toByteArray();
        }
        // obteniendo la inversion
        MaeInversion inversion = oMaeEstadoCuentas.get(0).getMaeInversion();
        // nombre de la hoja del excel
        workbook.setSheetName(0, inversion.getCInversion().trim());

        XSSFCell cell;

        // codigo
        XSSFRow rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("CÓDIGO:");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(inversion.getCInversion().trim());
        cell = rowHeader.createCell(6);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("FECHA INICIO:");
        cell = rowHeader.createCell(7);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(formatter.format(inversion.getFEmision()));
        // dni
        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("DNI:");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(inversion.getcPersonaId().getANroDocumento());
        cell = rowHeader.createCell(6);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("FECHA FIN:");
        rowHeader.createCell(7).setCellValue(formatter.format(inversion.getFVencimiento()));
        // persona
        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("APELLIDOS Y NOMBRES:");
        cell = rowHeader.createCell(2);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(inversion.getcPersonaId().getDApePat() + " "
                + inversion.getcPersonaId().getDApeMat() + " "
                + inversion.getcPersonaId().getDNombres());
        cell = rowHeader.createCell(6);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("TASA:");
        cell = rowHeader.createCell(7);
        cell.setCellStyle(headerStyle);
        cell.setCellValue((double) inversion.getPTasa() * 100 + "%");
        // direccion
        rowHeader = sheet.createRow((short) xLine++);
        cell = rowHeader.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("DIRECCIÓN:");
        if (inversion.getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist() != null) {
            cell = rowHeader.createCell(2);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(inversion.getcPersonaId().getMaeDireccionList().get(0).getADir1()
                    + "  " + inversion.getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoProv().trim()
                    + " - " + inversion.getcPersonaId().getMaeDireccionList().get(0).getMaeUbigeo().getDDUbigeoDist().trim());
        } else {
            rowHeader.createCell(1).setCellValue(" ");
        }
        cell = rowHeader.createCell(6);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("MONTO:");
        cell = rowHeader.createCell(7);
        cell.setCellStyle(headerStyle);
        cell.setCellValue((double) inversion.getIInversion());
        xLine = xLine + 1;
        rowHeader = sheet.createRow((short) xLine);

        String[] headers = new String[]{
            "N°", "Cuota", "Fecha Operación", "Descripción de la operación", "Abono", "Cago", "Saldo"
        };

        for (int i = 0; i < headers.length; ++i) {
            String header = headers[i];
            cell = rowHeader.createCell(i + 1);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(header);
        }
        xLine = xLine + 1;
        for (MaeEstadoCuenta estCta : oMaeEstadoCuentas) {

            rowHeader = sheet.createRow((short) xLine++);
            // nro correlativo
            cell = rowHeader.createCell(1);
            cell.setCellStyle(style);
            cell.setCellValue(estCta.getNsecuencia());
            // nro cuota
            cell = rowHeader.createCell(2);
            cell.setCellStyle(style);
            cell.setCellValue(estCta.getDcuota());
            // fec operacion
            cell = rowHeader.createCell(3);
            cell.setCellStyle(style);
            cell.setCellValue(formatter.format(estCta.getFoperacion()));
            // dec operacion
            cell = rowHeader.createCell(4);
            cell.setCellStyle(style);
            cell.setCellValue(estCta.getDdescriocion());
            // abono
            cell = rowHeader.createCell(5);
            cell.setCellStyle(style);
            cell.setCellValue((float) estCta.getIabono());
            // cargo
            cell = rowHeader.createCell(6);
            cell.setCellStyle(style);
            cell.setCellValue((float) estCta.getIcargo());
            // saldo
            cell = rowHeader.createCell(7);
            cell.setCellStyle(style);
            cell.setCellValue((float) estCta.getIsaldo());

        }
        workbook.write(byteOutputStream);
        
        POIXMLProperties xmlProps = workbook.getProperties();
        POIXMLProperties.CoreProperties coreProps =  xmlProps.getCoreProperties();
        coreProps.setCreator(maeReporte.getUserName());
        // grabando archivo para autidoria
        Date date = new Date();
        LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        String ruta;
        String uuid;
        uuid = oMaeEstadoCuentas.get(0).getMaeInversion().getCInversion().trim() + "_" + UUID.randomUUID().toString() + ".xlsx";
        ruta = "D:\\webcobranzas\\files\\account_status\\" + localDate.getYear()
                + "\\" + localDate.getMonthValue()
                + "\\" + localDate.getDayOfMonth()
                + "\\" + uuid;
        File fileO = new File(ruta);
        fileO.getParentFile().mkdirs();
        FileOutputStream fos = new FileOutputStream(fileO);
        byteOutputStream.writeTo(fos);
        fos.close();

        return byteOutputStream.toByteArray();

    }
}
