/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.archivo;

import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.UUID;
import javax.ejb.Stateless;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import pop.comun.dominio.MaeInversion;
import pop.comun.dominio.MaePersonaInmueble;
import pop.comun.dominio.MaeReporte;
import pop.comun.dominio.TabArchivo;

/**
 *
 * @author Jyoverar
 */
@Stateless(mappedName = "ejbArchivoDoc")
public class SessionArcDocumento implements IArcDocumento {

    @Override
    public byte[] prePlazo24H(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));
        String newline = "\n";
        int tipoD = maeReporte.getTipoDoc();
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();

        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "1.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        System.out.println("uuid = " + uuid);

        // 3) Convert XWPFDocument to Pdf
        //File fileO = new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\" + uuid + ".pdf");
        //fileO.getParentFile().mkdirs();
        //OutputStream out = new FileOutputStream(fileO);//new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\00011.pdf"));
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();

        document.write(byteOutputStream);
        //PdfConverter.getInstance().convert(document, byteOutputStream, options);

        //out.close();
        is.close();
        return byteOutputStream.toByteArray();

    }

    @Override
    public TabArchivo genPlazo24H(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));
        String newline = "\n";

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();
        TabArchivo oTabArchivo = new TabArchivo();

        String ruta;
        String name;
        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "1.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        //System.out.println("uuid = " + uuid);

        try {
            //name = maeInversion.getCcodigoIdent().trim() + "_PL24H_" + uuid + ".pdf";
            name = maeInversion.getCcodigoIdent().trim() + "_PL24H_" + uuid + ".docx";
            ruta = "D:\\webcobranzas\\files\\document\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "\\" + maeInversion.getCcodigoIdent() + "\\" + name;
            // 3) Convert XWPFDocument to Pdf
            File fileO = new File(ruta);
            fileO.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(fileO);
            //PdfConverter.getInstance().convert(document, out, options);
            document.write(out);
            out.close();
            //
            //oTabArchivo.setCtipoArcId("0003");
            oTabArchivo.setCtipoArcId("0001");
            oTabArchivo.setDnombreArc(name);
            oTabArchivo.setDruta(ruta);
            oTabArchivo.setBgenerado(true);
        } catch (XWPFConverterException | IOException e) {
            is.close();
            System.out.println("pop.webcobranzas.archivo.SessionArcDocumento.genPlazo24H() - " + e.getMessage());
        }

        is.close();
        return oTabArchivo;

    }

    @Override
    public byte[] preProtesto(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {
        int tipoD = maeReporte.getTipoDoc();
        Locale.setDefault(new Locale("es", "ES"));
        String newline = "\n";

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();

        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "3.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }

                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // reemplaza al Fondo
                            if (text != null && text.contains("$FONDO")) {
                                text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        System.out.println("uuid = " + uuid);

        // 3) Convert XWPFDocument to Pdf
        //File fileO = new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\" + uuid + ".pdf");
        //fileO.getParentFile().mkdirs();
        //OutputStream out = new FileOutputStream(fileO);//new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\00011.pdf"));
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();

        //PdfConverter.getInstance().convert(document, byteOutputStream, options);
        document.write(byteOutputStream);

        //out.close();
        is.close();
        return byteOutputStream.toByteArray();

    }

    @Override
    public TabArchivo genProtesto(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));
        String newline = "\n";

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();
        TabArchivo oTabArchivo = new TabArchivo();

        String ruta;
        String name;
        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "3.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // reemplaza al Fondo
                            if (text != null && text.contains("$FONDO")) {
                                text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        //System.out.println("uuid = " + uuid);

        try {
            //name = maeInversion.getCcodigoIdent().trim() + "_PROTESTO_" + uuid + ".pdf";
            name = maeInversion.getCcodigoIdent().trim() + "_PROTESTO_" + uuid + ".docx";
            ruta = "D:\\webcobranzas\\files\\document\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "\\" + maeInversion.getCcodigoIdent() + "\\" + name;

            // 3) Convert XWPFDocument to Pdf
            File fileO = new File(ruta);
            fileO.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(fileO);
            //PdfConverter.getInstance().convert(document, out, options);
            document.write(out);
            out.close();
            //
            //oTabArchivo.setCtipoArcId("0003");
            oTabArchivo.setCtipoArcId("0001");
            oTabArchivo.setDnombreArc(name);
            oTabArchivo.setDruta(ruta);
            oTabArchivo.setBgenerado(true);
        } catch (XWPFConverterException | IOException e) {
            is.close();
            System.out.println("pop.webcobranzas.archivo.SessionArcDocumento.genProtesto() - " + e.getMessage());
        }

        is.close();
        return oTabArchivo;

    }

    @Override
    public byte[] preJudicial(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));
        String newline = "\n";

        int tipoD = maeReporte.getTipoDoc();
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();

        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "5.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }

                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // reemplaza al Fondo
                            if (text != null && text.contains("$FONDO")) {
                                text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        System.out.println("uuid = " + uuid);

        // 3) Convert XWPFDocument to Pdf
        //File fileO = new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\" + uuid + ".pdf");
        //fileO.getParentFile().mkdirs();
        //OutputStream out = new FileOutputStream(fileO);//new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\00011.pdf"));
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();

        // PdfConverter.getInstance().convert(document, byteOutputStream, options);
        document.write(byteOutputStream);

        //out.close();
        is.close();
        return byteOutputStream.toByteArray();

    }

    @Override
    public TabArchivo genPreJudicial(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));
        String newline = "\n";

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();
        TabArchivo oTabArchivo = new TabArchivo();

        String ruta;
        String name;
        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "5.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // reemplaza al Fondo
                            if (text != null && text.contains("$FONDO")) {
                                text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }

                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        //System.out.println("uuid = " + uuid);

        try {
            //name = maeInversion.getCcodigoIdent().trim() + "_PREJUDICIAL_" + uuid + ".pdf";
            name = maeInversion.getCcodigoIdent().trim() + "_PREJUDICIAL_" + uuid + ".docx";
            ruta = "D:\\webcobranzas\\files\\document\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "\\" + maeInversion.getCcodigoIdent() + "\\" + name;

            // 3) Convert XWPFDocument to Pdf
            File fileO = new File(ruta);
            fileO.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(fileO);
            //PdfConverter.getInstance().convert(document, out, options);
            document.write(out);
            out.close();
            //
            //oTabArchivo.setCtipoArcId("0003");
            oTabArchivo.setCtipoArcId("0001");
            oTabArchivo.setDnombreArc(name);
            oTabArchivo.setDruta(ruta);
            oTabArchivo.setBgenerado(true);
        } catch (XWPFConverterException | IOException e) {
            is.close();
            System.out.println("pop.webcobranzas.archivo.SessionArcDocumento.genProtesto() - " + e.getMessage());
        }

        is.close();
        return oTabArchivo;

    }

    @Override
    public byte[] preUltAviso(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));
        String newline = "\n";

        int tipoD = maeReporte.getTipoDoc();
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();

        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "7.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplaza al Fondo
                    if (text != null && text.contains("$FONDO")) {
                        text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // reemplaza al Fondo
                            if (text != null && text.contains("$FONDO")) {
                                text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        System.out.println("uuid = " + uuid);

        // 3) Convert XWPFDocument to Pdf
        //File fileO = new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\" + uuid + ".pdf");
        //fileO.getParentFile().mkdirs();
        //OutputStream out = new FileOutputStream(fileO);//new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\00011.pdf"));
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();

        //PdfConverter.getInstance().convert(document, byteOutputStream, options);
        document.write(byteOutputStream);

        //out.close();
        is.close();
        return byteOutputStream.toByteArray();
    }

    @Override
    public TabArchivo genUltAviso(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));
        String newline = "\n";

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();
        TabArchivo oTabArchivo = new TabArchivo();

        String ruta;
        String name;
        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "7.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplaza al Fondo
                    if (text != null && text.contains("$FONDO")) {
                        text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                        r.setText(text, 0);
                    }
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Fondo
                            if (text != null && text.contains("$FONDO")) {
                                text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        //System.out.println("uuid = " + uuid);

        try {
            //name = maeInversion.getCcodigoIdent().trim() + "_ULTAVI_" + uuid + ".pdf";
            name = maeInversion.getCcodigoIdent().trim() + "_ULTAVISO_" + uuid + ".docx";
            ruta = "D:\\webcobranzas\\files\\document\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "\\" + maeInversion.getCcodigoIdent() + "\\" + name;
            // 3) Convert XWPFDocument to Pdf
            File fileO = new File(ruta);
            fileO.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(fileO);
            //PdfConverter.getInstance().convert(document, out, options);
            document.write(out);
            out.close();
            //
            //oTabArchivo.setCtipoArcId("0003");
            oTabArchivo.setCtipoArcId("0001");
            oTabArchivo.setDnombreArc(name);
            oTabArchivo.setDruta(ruta);
            oTabArchivo.setBgenerado(true);
        } catch (XWPFConverterException | IOException e) {
            is.close();
            System.out.println("pop.webcobranzas.archivo.SessionArcDocumento.genUltAviso() - " + e.getMessage());
        }

        is.close();
        return oTabArchivo;
    }

    @Override
    public byte[] preNegExtJudicial(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));

        int tipoD = maeReporte.getTipoDoc();
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");

        Locale loc = new Locale("en", "US");
        NumberFormat numberFormatter = NumberFormat.getInstance(loc);
        numberFormatter.setMaximumFractionDigits(2);

        // para los formatos de los numero y fechas
        //DecimalFormat formatterNum = new DecimalFormat("###,###,##0.00");
        Date fecha = maeInversion.getfIniBusq();

        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "9.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }

                    // reemplaza al meses pactados
                    if (text != null && text.contains("$MESES")) {
                        text = text.replace("$MESES", Integer.toString((int) maeInversion.getNMeses()));
                        r.setText(text, 0);
                    }

                    // reemplaza la cuota de moda
                    if (text != null && text.contains("$ICUOTAPAC")) {
                        text = text.replace("$ICUOTAPAC", numberFormatter.format(maeInversion.getICuota()));
                        r.setText(text, 0);
                    }

                    // reemplaza al FECHACUOTA
                    if (text != null && text.contains("GT")) {
                        text = text.replace("GT", formatter.format(fecha));
                        r.setText(text, 0);
                    }

                    // reemplaza al saldo deudor
                    if (text != null && text.contains("$SALDODEUDOR")) {

                        text = text.replace("$SALDODEUDOR", numberFormatter.format(maeInversion.getRepSaldoDeudor().getNtotDebe() + maeInversion.getRepSaldoDeudor().getIgastLegalFut() + maeInversion.getRepSaldoDeudor().getIgastAdmin()));
                        r.setText(text, 0);
                    }
                    //formatterNum.format(oSaldoDeudor.getNtotDebe() + oSaldoDeudor.getIgastLegalFut() + oSaldoDeudor.getIgastAdmin())

                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // firma del usuario
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                            // reemplaza al meses pactados
                            if (text != null && text.contains("$MESES")) {
                                text = text.replace("$MESES", Integer.toString((int) maeInversion.getNMeses()));
                                r.setText(text, 0);
                            }
                            // reemplaza la cuota de moda
                            if (text != null && text.contains("$ICUOTAPAC")) {
                                text = text.replace("$ICUOTAPAC", String.valueOf(maeInversion.getICuota()));
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        System.out.println("uuid = " + uuid);

        // 3) Convert XWPFDocument to Pdf
        //File fileO = new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\" + uuid + ".pdf");
        //fileO.getParentFile().mkdirs();
        //OutputStream out = new FileOutputStream(fileO);//new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\00011.pdf"));
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();

        document.write(byteOutputStream);
        //PdfConverter.getInstance().convert(document, byteOutputStream, options);

        //out.close();
        is.close();
        return byteOutputStream.toByteArray();

    }

    @Override
    public TabArchivo genNegExtJudicial(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();
        TabArchivo oTabArchivo = new TabArchivo();

        Locale loc = new Locale("en", "US");
        NumberFormat numberFormatter = NumberFormat.getInstance(loc);
        numberFormatter.setMaximumFractionDigits(2);

        String ruta;
        String name;
        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "9.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }

                    // reemplaza al meses pactados
                    if (text != null && text.contains("$MESES")) {
                        text = text.replace("$MESES", Integer.toString((int) maeInversion.getNMeses()));
                        r.setText(text, 0);
                    }

                    // reemplaza la cuota de moda
                    if (text != null && text.contains("$ICUOTAPAC")) {
                        text = text.replace("$ICUOTAPAC", numberFormatter.format(maeInversion.getICuota()));
                        r.setText(text, 0);
                    }

                    // reemplaza al FECHACUOTA
                    if (text != null && text.contains("GT")) {
                        text = text.replace("GT", formatter.format(fecha));
                        r.setText(text, 0);
                    }

                    // reemplaza al saldo deudor
                    if (text != null && text.contains("$SALDODEUDOR")) {

                        text = text.replace("$SALDODEUDOR", numberFormatter.format(maeInversion.getRepSaldoDeudor().getNtotDebe() + maeInversion.getRepSaldoDeudor().getIgastLegalFut() + maeInversion.getRepSaldoDeudor().getIgastAdmin()));
                        r.setText(text, 0);
                    }
                    //formatterNum.format(oSaldoDeudor.getNtotDebe() + oSaldoDeudor.getIgastLegalFut() + oSaldoDeudor.getIgastAdmin())

                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // firma del usuario
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                            // reemplaza al meses pactados
                            if (text != null && text.contains("$MESES")) {
                                text = text.replace("$MESES", Integer.toString((int) maeInversion.getNMeses()));
                                r.setText(text, 0);
                            }
                            // reemplaza la cuota de moda
                            if (text != null && text.contains("$ICUOTAPAC")) {
                                text = text.replace("$ICUOTAPAC", String.valueOf(maeInversion.getICuota()));
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        //System.out.println("uuid = " + uuid);

        try {
            //name = maeInversion.getCcodigoIdent().trim() + "_PL24H_" + uuid + ".pdf";
            name = maeInversion.getCcodigoIdent().trim() + "_NegExtJud_" + uuid + ".docx";
            ruta = "D:\\webcobranzas\\files\\document\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "\\" + maeInversion.getCcodigoIdent() + "\\" + name;
            // 3) Convert XWPFDocument to Pdf
            File fileO = new File(ruta);
            fileO.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(fileO);
            //PdfConverter.getInstance().convert(document, out, options);
            document.write(out);
            out.close();
            //
            //oTabArchivo.setCtipoArcId("0003");
            oTabArchivo.setCtipoArcId("0001");
            oTabArchivo.setDnombreArc(name);
            oTabArchivo.setDruta(ruta);
            oTabArchivo.setBgenerado(true);
        } catch (XWPFConverterException | IOException e) {
            is.close();
            System.out.println("pop.webcobranzas.archivo.SessionArcDocumento.genNegExtJudicial() - " + e.getMessage());
        }

        is.close();
        return oTabArchivo;

    }

    @Override
    public byte[] preNegPreJudicial(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {
        Locale.setDefault(new Locale("es", "ES"));

        int tipoD = maeReporte.getTipoDoc();
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");

        Locale loc = new Locale("en", "US");
        NumberFormat numberFormatter = NumberFormat.getInstance(loc);
        numberFormatter.setMaximumFractionDigits(2);

        // para los formatos de los numero y fechas
        //DecimalFormat formatterNum = new DecimalFormat("###,###,##0.00");
        Date fecha = maeInversion.getfIniBusq();

        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "10.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }

                    // reemplaza al CUOTAS
                    if (text != null && text.contains("$CUOTAS")) {
                        text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                        r.setText(text, 0);
                    }

                    // reemplaza al meses pactados
                    if (text != null && text.contains("$MESES")) {
                        text = text.replace("$MESES", Integer.toString((int) maeInversion.getNMeses()));
                        r.setText(text, 0);
                    }

                    // reemplaza el estado de cuenta a la fecha 
                    if (text != null && text.contains("$ICUOTAPAC")) {
                        text = text.replace("$ICUOTAPAC", numberFormatter.format(maeInversion.getICuota()));
                        r.setText(text, 0);
                    }

                    // reemplaza la cuota de moda
                    if (text != null && text.contains("DEUDA")) {
                        text = text.replace("DEUDA", numberFormatter.format(((-1) * (float) maeInversion.getMaeCuotaPagoList().get(0).getiPendiente())));
                        r.setText(text, 0);
                    }

                    // reemplaza al FECHACUOTA
                    if (text != null && text.contains("GT")) {
                        text = text.replace("GT", formatter.format(fecha));
                        r.setText(text, 0);
                    }

                    // reemplaza al saldo deudor
                    if (text != null && text.contains("$SALDODEUDOR")) {

                        text = text.replace("$SALDODEUDOR", numberFormatter.format(maeInversion.getRepSaldoDeudor().getNtotDebe() + maeInversion.getRepSaldoDeudor().getIgastLegalFut() + maeInversion.getRepSaldoDeudor().getIgastAdmin()));
                        r.setText(text, 0);
                    }
                    // reemplaza al DEMANDA
                    if (text != null && text.contains("DEMANDA")) {
                        text = text.replace("DEMANDA", formatter.format(maeInversion.getfFinBusq()));
                        r.setText(text, 0);
                    }
                    //
                    // reemplaza al DEMANDA
                    if (text != null && text.contains("FIN")) {
                        text = text.replace("FIN", formatter.format(maeInversion.getRepSaldoDeudor().getMaeInversion().getFVencimiento()));
                        r.setText(text, 0);
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al DEMANDA
                            if (text != null && text.contains("DEMANDA")) {
                                text = text.replace("DEMANDA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // firma del usuario
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                            // reemplaza al meses pactados
                            if (text != null && text.contains("$MESES")) {
                                text = text.replace("$MESES", Integer.toString((int) maeInversion.getNMeses()));
                                r.setText(text, 0);
                            }
                            // reemplaza la cuota de moda
                            if (text != null && text.contains("$ICUOTAPAC")) {
                                text = text.replace("$ICUOTAPAC", String.valueOf(maeInversion.getICuota()));
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        System.out.println("uuid = " + uuid);

        // 3) Convert XWPFDocument to Pdf
        //File fileO = new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\" + uuid + ".pdf");
        //fileO.getParentFile().mkdirs();
        //OutputStream out = new FileOutputStream(fileO);//new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\00011.pdf"));
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();

        document.write(byteOutputStream);
        //PdfConverter.getInstance().convert(document, byteOutputStream, options);

        //out.close();
        is.close();
        return byteOutputStream.toByteArray();

    }

    @Override
    public TabArchivo genNegPreJudicial(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {
        Locale.setDefault(new Locale("es", "ES"));

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();
        TabArchivo oTabArchivo = new TabArchivo();

        Locale loc = new Locale("en", "US");
        NumberFormat numberFormatter = NumberFormat.getInstance(loc);
        numberFormatter.setMaximumFractionDigits(2);

        String ruta;
        String name;
        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "10.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }

                    // reemplaza al CUOTAS
                    if (text != null && text.contains("$CUOTAS")) {
                        text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                        r.setText(text, 0);
                    }

                    // reemplaza al meses pactados
                    if (text != null && text.contains("$MESES")) {
                        text = text.replace("$MESES", Integer.toString((int) maeInversion.getNMeses()));
                        r.setText(text, 0);
                    }

                    // reemplaza la cuota de moda
                    if (text != null && text.contains("$ICUOTAPAC")) {
                        text = text.replace("$ICUOTAPAC", numberFormatter.format(maeInversion.getICuota()));
                        r.setText(text, 0);
                    }

                    // reemplaza al FECHACUOTA
                    if (text != null && text.contains("GT")) {
                        text = text.replace("GT", formatter.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al saldo deudor
                    if (text != null && text.contains("$SALDODEUDOR")) {

                        text = text.replace("$SALDODEUDOR", numberFormatter.format(maeInversion.getRepSaldoDeudor().getNtotDebe() + maeInversion.getRepSaldoDeudor().getIgastLegalFut() + maeInversion.getRepSaldoDeudor().getIgastAdmin()));
                        r.setText(text, 0);
                    }
                    // reemplaza al DEMANDA
                    if (text != null && text.contains("DEMANDA")) {
                        text = text.replace("DEMANDA", formatter.format(maeInversion.getfFinBusq()));
                        r.setText(text, 0);
                    }
                    
                    // reemplaza la cuota de moda
                    if (text != null && text.contains("DEUDA")) {
                        text = text.replace("DEUDA", numberFormatter.format(((-1) * (float) maeInversion.getMaeCuotaPagoList().get(0).getiPendiente())));
                        r.setText(text, 0);
                    }
                    
                    //
                    // reemplaza al DEMANDA
                    if (text != null && text.contains("FIN")) {
                        text = text.replace("FIN", formatter.format(maeInversion.getRepSaldoDeudor().getMaeInversion().getFVencimiento()));
                        r.setText(text, 0);
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                String repre = "";
                                boolean flag = true;
                                for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                                    repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                                    if (flag) {
                                        text = text.replace("$CLIENTE", repre);
                                        r.setText(text, 0);
                                        flag = false;
                                    } else {
                                        text = repre;
                                        r.addBreak();
                                        r.setText(text);
                                    }
                                }
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al DEMANDA
                            if (text != null && text.contains("DEMANDA")) {
                                text = text.replace("DEMANDA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // firma del usuario
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                            // reemplaza al meses pactados
                            if (text != null && text.contains("$MESES")) {
                                text = text.replace("$MESES", Integer.toString((int) maeInversion.getNMeses()));
                                r.setText(text, 0);
                            }
                            // reemplaza la cuota de moda
                            if (text != null && text.contains("$ICUOTAPAC")) {
                                text = text.replace("$ICUOTAPAC", String.valueOf(maeInversion.getICuota()));
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        //System.out.println("uuid = " + uuid);

        try {
            //name = maeInversion.getCcodigoIdent().trim() + "_PL24H_" + uuid + ".pdf";
            name = maeInversion.getCcodigoIdent().trim() + "_NegPreJud_" + uuid + ".docx";
            ruta = "D:\\webcobranzas\\files\\document\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "\\" + maeInversion.getCcodigoIdent() + "\\" + name;
            // 3) Convert XWPFDocument to Pdf
            File fileO = new File(ruta);
            fileO.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(fileO);
            //PdfConverter.getInstance().convert(document, out, options);
            document.write(out);
            out.close();
            //
            //oTabArchivo.setCtipoArcId("0003");
            oTabArchivo.setCtipoArcId("0001");
            oTabArchivo.setDnombreArc(name);
            oTabArchivo.setDruta(ruta);
            oTabArchivo.setBgenerado(true);
        } catch (XWPFConverterException | IOException e) {
            is.close();
            System.out.println("pop.webcobranzas.archivo.SessionArcDocumento.genNegExtJudicial() - " + e.getMessage());
        }

        is.close();
        return oTabArchivo;
    }

    @Override
    public byte[] imprimirArchivo(TabArchivo tabArchivo) throws Exception {

        File outputFile = new File(tabArchivo.getDruta());
        Path p1 = Paths.get(tabArchivo.getDruta());
        byte[] demBytes = Files.readAllBytes(p1);

//        try {
//            OutputStream outputStream = new FileOutputStream(outputFile);
//            outputStream.write(demBytes);
//            outputStream.close();
//        } catch (Exception e) {
//            
//        } 
        return demBytes;
    }

    @Override
    public TabArchivo guardarArchivo(MaeInversion maeInversion, TabArchivo archivo) throws Exception {
        String ruta;
        String uuid = UUID.randomUUID().toString();
        String name = maeInversion.getCcodigoIdent().trim() + "_FILE_" + uuid;// archivo.getDnombreArc();
        String tipoFile;
        String ctipoFile = "";

        tipoFile = archivo.getDnombreArc().substring(archivo.getDnombreArc().lastIndexOf(".") + 1).toUpperCase();

        switch (tipoFile) {
            case "DOC":
                ctipoFile = "0001";
                break;
            case "DOCX":
                ctipoFile = "0001";
                break;
            case "XLS":
                ctipoFile = "0002";
                break;
            case "XLSX":
                ctipoFile = "0002";
                break;
            case "PDF":
                ctipoFile = "0003";
                break;
            case "JPG":
                ctipoFile = "0004";
                break;
            case "PNG":
                ctipoFile = "0004";
                break;
            case "BMP":
                ctipoFile = "0004";
                break;
        }

        TabArchivo oTabArchivo = new TabArchivo();

        ruta = "D:\\webcobranzas\\files\\document\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "\\" + maeInversion.getCcodigoIdent() + "\\" + name + "." + tipoFile;

        File fileO = new File(ruta);
        fileO.getParentFile().mkdirs();
        OutputStream out = new FileOutputStream(fileO);
        out.write(archivo.getDdatos());
        out.close();
        //

        System.out.println("   tipoFile --> " + tipoFile + " " + ctipoFile);

        oTabArchivo.setDnombreArc(name + '.' + tipoFile);
        oTabArchivo.setDruta(ruta);
        oTabArchivo.setCtipoArcId(ctipoFile);
        oTabArchivo.setBgenerado(true);

        return oTabArchivo;

    }

    // ------------------------- judiciales
    @Override
    public byte[] preCN(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();

        String newline = "\n";

        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "2.docx");
        XWPFDocument document = new XWPFDocument(is);

        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al codigo
                    if (text != null && text.contains("$CODIGO")) {
                        text = text.replace("$CODIGO", maeInversion.getCcodigoIdent().trim());
                        r.setText(text, 0);
                    }
                    // reemplaza a los representantes
                    if (text != null && text.contains("$REPRESENTANTES")) {
                        String repre = "";
                        boolean flag = true;
                        for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                            repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                            if (flag) {
                                text = text.replace("$REPRESENTANTES", repre);
                                r.setText(text, 0);
                                flag = false;
                            } else {
                                text = repre;
                                r.addBreak();
                                r.setText(text);
                            }
                        }
                    }
                    // reemplaza al domicilio
                    if (text != null && text.contains("$DOMICILIO")) {
                        text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getDDir1() + " "
                                + maeInversion.getMaeInmueble().getMaeUbigeo().getDUbigeo().trim() + " - "
                                + maeInversion.getMaeInmueble().getMaeUbigeo().getMaeUbigeo().getDUbigeo().trim());
                        r.setText(text, 0);
                    }
                    // reemplaza al referencia
                    if (text != null && text.contains("$REFERENCIA")) {
                        if (maeInversion.getMaeInmueble().getDdir3() != null) {
                            text = text.replace("$REFERENCIA", maeInversion.getMaeInmueble().getDdir3());
                        } else {
                            text = text.replace("$REFERENCIA", "");
                        }

                        r.setText(text, 0);
                    }
                    // reemplaza al asiento
                    if (text != null && text.contains("$ASIENTO")) {
                        text = text.replace("$ASIENTO", maeInversion.getMaeInmueble().getMaeHipoteca().getCasientoId());
                        r.setText(text, 0);
                    }
                    // reemplaza al partida
                    if (text != null && text.contains("$PARTIDAELEC")) {
                        text = text.replace("$PARTIDAELEC", maeInversion.getMaeInmueble().getMaeHipoteca().getCpartidaElecId());
                        r.setText(text, 0);
                    }
                    // reemplaza al sede
                    if (text != null && text.contains("$SEDE")) {
                        text = text.replace("$SEDE", maeInversion.getMaeInmueble().getMaeHipoteca().getCsede().getDdescCorta());
                        r.setText(text, 0);
                    }
                    // reemplaza al FESCRITURA 
                    if (text != null && text.contains("$FESCRITURA")) {
                        text = text.replace("$FESCRITURA",
                                formatterDia.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFasiento()) + " de " + formatterMes.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFasiento()) + " del " + formatterAnio.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFasiento()));
                        r.setText(text, 0);
                    }
                    // reemplaza al notario 
                    if (text != null && text.contains("$NOTARIO")) {
                        text = text.replace("$NOTARIO", maeInversion.getMaeInmueble().getMaeHipoteca().getDnomNotaria());
                        r.setText(text, 0);
                    }
                    // reemplaza al TCHN  
                    if (text != null && text.contains("$TCHN")) {
                        text = text.replace("$TCHN", maeInversion.getMaeInmueble().getMaeHipoteca().getCtchnReal());
                        r.setText(text, 0);
                    }
                    // reemplaza al FTCHN  
                    if (text != null && text.contains("$FTCHN")) {
                        text = text.replace("$FTCHN",
                                formatterDia.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFemisionTchn()) + " de " + formatterMes.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFemisionTchn()) + " del " + formatterAnio.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFemisionTchn()));
                        r.setText(text, 0);
                    }
                    // reemplaza al ATCHN   
                    if (text != null && text.contains("$ATCHN")) {
                        text = text.replace("$ATCHN", maeInversion.getMaeInmueble().getMaeHipoteca().getCasientoTchn());
                        r.setText(text, 0);
                    }
                    // reemplaza al $SEDEB   
                    if (text != null && text.contains("#sb")) {
                        text = text.replace("#sb", maeInversion.getMaeInmueble().getMaeHipoteca().getCsedeTchn().getDdescCorta());
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getcUsuarioAdd());
                        r.setText(text, 0);
                    }
                    // cuotas que debe
                    if (text != null && text.contains("$CUOTASC")) {
                        text = text.replace("$CUOTASC", Integer.toString((int) maeInversion.getNCuotasAtrasadas()) + "-C");
                        r.setText(text, 0);
                    }
                    // reemplaza al Fondo
                    if (text != null && text.contains("$FONDO")) {
                        text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }

                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            System.out.println(" B ------> " + text);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                text = text.replace("$CLIENTE", maeInversion.getcPersonaId().getDApePat() + " " + maeInversion.getcPersonaId().getDApeMat() + " " + maeInversion.getcPersonaId().getDNombres());
                                r.setText(text, 0);
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }
                            // reemplaza al Fondo
                            if (text != null && text.contains("$FONDO")) {
                                text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        System.out.println("uuid = " + uuid);

        // 3) Convert XWPFDocument to Pdf
        //File fileO = new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\" + uuid + ".pdf");
        //fileO.getParentFile().mkdirs();
        //OutputStream out = new FileOutputStream(fileO);//new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\00011.pdf"));
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();

        //PdfConverter.getInstance().convert(document, byteOutputStream, options);
        document.write(byteOutputStream);

        //out.close();
        is.close();
        return byteOutputStream.toByteArray();
    }

    @Override
    public TabArchivo genCN(MaeInversion maeInversion, MaeReporte maeReporte) throws Exception {

        Locale.setDefault(new Locale("es", "ES"));

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = maeInversion.getfIniBusq();
        TabArchivo oTabArchivo = new TabArchivo();

        String ruta;
        String name;

        String newline = "\n";

        //fecha = formatter.parse("07/04/2017");
        // 1) Load DOCX into XWPFDocument
        InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "2.docx");
        XWPFDocument document = new XWPFDocument(is);
        // dato sueltos
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);

                    System.err.println("------> " + text);
                    // reemplaza la fecha
                    if (text != null && text.contains("$FECHA")) {
                        text = text.replace("$FECHA", formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));
                        r.setText(text, 0);
                    }
                    // reemplaza al codigo
                    if (text != null && text.contains("$CODIGO")) {
                        text = text.replace("$CODIGO", maeInversion.getCcodigoIdent().trim());
                        r.setText(text, 0);
                    }
                    // reemplaza a los representantes
                    if (text != null && text.contains("$REPRESENTANTES")) {
                        String repre = "";
                        boolean flag = true;
                        for (MaePersonaInmueble perinm : maeInversion.getMaeInmueble().getMaePersonaInmuebleList()) {
                            repre = perinm.getMaePersona().getDNombres() + " " + perinm.getMaePersona().getDApePat() + " " + perinm.getMaePersona().getDApeMat();
                            if (flag) {
                                text = text.replace("$REPRESENTANTES", repre);
                                r.setText(text, 0);
                                flag = false;
                            } else {
                                text = repre;
                                r.addBreak();
                                r.setText(text);
                            }
                        }
                    }
                    // reemplaza al domicilio
                    if (text != null && text.contains("$DOMICILIO")) {
                        text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getDDir1() + " "
                                + maeInversion.getMaeInmueble().getMaeUbigeo().getDUbigeo().trim() + " - "
                                + maeInversion.getMaeInmueble().getMaeUbigeo().getMaeUbigeo().getDUbigeo().trim());
                        r.setText(text, 0);
                    }
                    // reemplaza al referencia
                    if (text != null && text.contains("$REFERENCIA")) {
                        if (maeInversion.getMaeInmueble().getDdir3() != null) {
                            text = text.replace("$REFERENCIA", maeInversion.getMaeInmueble().getDdir3());
                        } else {
                            text = text.replace("$REFERENCIA", "");
                        }

                        r.setText(text, 0);
                    }
                    // reemplaza al asiento
                    if (text != null && text.contains("$ASIENTO")) {
                        text = text.replace("$ASIENTO", maeInversion.getMaeInmueble().getMaeHipoteca().getCasientoId());
                        r.setText(text, 0);
                    }
                    // reemplaza al partida
                    if (text != null && text.contains("$PARTIDAELEC")) {
                        text = text.replace("$PARTIDAELEC", maeInversion.getMaeInmueble().getMaeHipoteca().getCpartidaElecId());
                        r.setText(text, 0);
                    }
                    // reemplaza al sede
                    if (text != null && text.contains("$SEDE")) {
                        text = text.replace("$SEDE", maeInversion.getMaeInmueble().getMaeHipoteca().getCsede().getDdescCorta());
                        r.setText(text, 0);
                    }
                    // reemplaza al FESCRITURA 
                    if (text != null && text.contains("$FESCRITURA")) {
                        text = text.replace("$FESCRITURA",
                                formatterDia.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFasiento()) + " de " + formatterMes.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFasiento()) + " del " + formatterAnio.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFasiento()));
                        r.setText(text, 0);
                    }
                    // reemplaza al notario 
                    if (text != null && text.contains("$NOTARIO")) {
                        text = text.replace("$NOTARIO", maeInversion.getMaeInmueble().getMaeHipoteca().getDnomNotaria());
                        r.setText(text, 0);
                    }
                    // reemplaza al TCHN  
                    if (text != null && text.contains("$TCHN")) {
                        text = text.replace("$TCHN", maeInversion.getMaeInmueble().getMaeHipoteca().getCtchnReal());
                        r.setText(text, 0);
                    }
                    // reemplaza al FTCHN  
                    if (text != null && text.contains("$FTCHN")) {
                        text = text.replace("$FTCHN",
                                formatterDia.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFemisionTchn()) + " de " + formatterMes.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFemisionTchn()) + " del " + formatterAnio.format(maeInversion.getMaeInmueble().getMaeHipoteca().getFemisionTchn()));
                        r.setText(text, 0);
                    }
                    // reemplaza al ATCHN   
                    if (text != null && text.contains("$ATCHN")) {
                        text = text.replace("$ATCHN", maeInversion.getMaeInmueble().getMaeHipoteca().getCasientoTchn());
                        r.setText(text, 0);
                    }
                    // reemplaza al $SEDEB   
                    if (text != null && text.contains("#sb")) {
                        text = text.replace("#sb", maeInversion.getMaeInmueble().getMaeHipoteca().getCsedeTchn().getDdescCorta());
                        r.setText(text, 0);
                    }
                    // reemplaza al Usuario
                    if (text != null && text.contains("$USUARIO")) {
                        text = text.replace("$USUARIO", maeReporte.getcUsuarioAdd());
                        r.setText(text, 0);
                    }

                    // cuotas que debe
                    if (text != null && text.contains("$CUOTASC")) {
                        text = text.replace("$CUOTASC", Integer.toString((int) maeInversion.getNCuotasAtrasadas()) + "-C");
                        r.setText(text, 0);
                    }
                    // reemplaza al Fondo
                    if (text != null && text.contains("$FONDO")) {
                        text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                        r.setText(text, 0);
                    }
                    // reemplazar la firma
                    if (text != null && text.contains("$FIRMA")) {
                        text = text.replace("$FIRMA", "");
                        r.setText(text, 0);
                        String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                        XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                        System.out.println(picture); //XWPFPicture is added
                        System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            // reemplaza al cliente
                            if (text != null && text.contains("$CLIENTE")) {
                                text = text.replace("$CLIENTE", maeInversion.getcPersonaId().getDApePat() + " " + maeInversion.getcPersonaId().getDApeMat() + " " + maeInversion.getcPersonaId().getDNombres());
                                r.setText(text, 0);
                            }
                            // reemplaza al domicilio
                            if (text != null && text.contains("$DOMICILIO")) {
                                text = text.replace("$DOMICILIO", maeInversion.getMaeInmueble().getADir1());
                                r.setText(text, 0);
                            }
                            // reemplaza al distrito
                            if (text != null && text.contains("$DISTRITO")) {
                                text = text.replace("$DISTRITO", maeInversion.getMaeInmueble().getMaeUbigeo().getDDUbigeoDist());
                                r.setText(text, 0);
                            }
                            // reemplaza al codigo
                            if (text != null && text.contains("$CODIGO")) {
                                text = text.replace("$CODIGO", maeInversion.getCInversion());
                                r.setText(text, 0);
                            }
                            // reemplaza al CUOTAS
                            if (text != null && text.contains("$CUOTAS")) {
                                text = text.replace("$CUOTAS", Integer.toString((int) maeInversion.getNCuotasAtrasadas()));
                                r.setText(text, 0);
                            }
                            // reemplaza al FECHACUOTA
                            if (text != null && text.contains("$FECHACUOTA")) {
                                text = text.replace("$FECHACUOTA", formatter.format(maeInversion.getfFinBusq()));
                                r.setText(text, 0);
                            }
                            // reemplaza al Usuario
                            if (text != null && text.contains("$USUARIO")) {
                                text = text.replace("$USUARIO", maeReporte.getDusuarioNombres() + " " + maeReporte.getDusuarioApellidos());
                                r.setText(text, 0);
                            }

                            // reemplaza al Fondo
                            if (text != null && text.contains("$FONDO")) {
                                text = text.replace("$FONDO", maeInversion.getMaeFondo().getDFondo());
                                r.setText(text, 0);
                            }

                            if (text != null && text.contains("$FIRMA")) {
                                text = text.replace("$FIRMA", "");
                                r.setText(text, 0);
                                String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\firmas\\" + maeReporte.getcUsuarioAdd() + ".png";
                                XWPFPicture picture = r.addPicture(new FileInputStream(imgFile), XWPFDocument.PICTURE_TYPE_PNG, imgFile, Units.toEMU(100), Units.toEMU(60));

                                System.out.println(picture); //XWPFPicture is added
                                System.out.println(picture.getPictureData()); //but without access to XWPFPictureData (no blipID)
                            }
                        }
                    }
                }
            }
        }
        // 2) Prepare Pdf options 
        //https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
        PdfOptions options = PdfOptions.create().fontEncoding("windows-1252");
        String uuid = UUID.randomUUID().toString();
        System.out.println("uuid = " + uuid);

        try {
            //name = maeInversion.getCcodigoIdent().trim() + "_CN_" + uuid + ".pdf";
            name = maeInversion.getCcodigoIdent().trim() + "_CN_" + uuid + ".docx";
            ruta = "D:\\webcobranzas\\files\\document\\" + maeInversion.getMaeFondo().getCFondoId().trim() + "\\" + maeInversion.getCcodigoIdent() + "\\" + name;

            // 3) Convert XWPFDocument to Pdf
            File fileO = new File(ruta);
            fileO.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(fileO);
            //PdfConverter.getInstance().convert(document, out, options);
            document.write(out);
            out.close();
            //
            oTabArchivo.setCtipoArcId("0001");
            oTabArchivo.setDnombreArc(name);
            oTabArchivo.setDruta(ruta);
            oTabArchivo.setBgenerado(true);
        } catch (XWPFConverterException | IOException e) {
            is.close();
            System.out.println("pop.webcobranzas.archivo.SessionArcDocumento.genCN() - " + e.getMessage());
        }

        is.close();
        return oTabArchivo;

    }

}
