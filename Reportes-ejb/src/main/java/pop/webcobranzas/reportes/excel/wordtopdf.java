/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.reportes.excel;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.UUID;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author Jyoverar
 */
public class wordtopdf {

    public static String readDocFile(String fileName) {

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        SimpleDateFormat formatterDia = new SimpleDateFormat("dd");
        SimpleDateFormat formatterMes = new SimpleDateFormat("MMMM");
        SimpleDateFormat formatterAnio = new SimpleDateFormat("yyyy");
        Date fecha = null;

        try {
            fecha = formatter.parse("07/04/2017");
        } catch (ParseException ex) {
            Logger.getLogger(wordtopdf.class.getName()).log(Level.SEVERE, null, ex);
        }

        System.out.println(" -> " + formatterDia.format(fecha) + " de " + formatterMes.format(fecha) + " del " + formatterAnio.format(fecha));

        String output = "";
        try {

            long start = System.currentTimeMillis();

            // 1) Load DOCX into XWPFDocument
            InputStream is = new FileInputStream(new File(fileName));
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
                        if (text != null && text.contains("$FIRMA")) {
                            text = text.replace("$FIRMA", "");
                            r.setText(text, 0);
                            String imgFile = "C:\\pop\\webcobranzas\\resources\\template\\JLEIVA.png";
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
                                    text = text.replace("$CLIENTE", "JHON ALEX YOVERA RAMOS");
                                    r.setText(text, 0);
                                }
                                // reemplaza al domicilio
                                if (text != null && text.contains("$DOMICILIO")) {
                                    text = text.replace("$DOMICILIO", "SAN JUAN BAUTISLA ASLD JKALSD AKJSD LAKSHDB AHSBD JAHSDJHASJD HA SD- ASDJA BHSDJ ASD- SD° ASD ");
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
            File fileO = new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\" + uuid + ".pdf");
            fileO.getParentFile().mkdirs();
            OutputStream out = new FileOutputStream(fileO);//new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\0002\\00011.pdf"));
            PdfConverter.getInstance().convert(document, out, options);

            out.close();

            is.close();

            System.err.println("Generate pdf/HelloWorld.pdf with "
                    + (System.currentTimeMillis() - start) + "ms");

//            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return output;
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        wordtopdf.readDocFile("C:\\pop\\webcobranzas\\resources\\template\\00011_2.docx");
    }

}
