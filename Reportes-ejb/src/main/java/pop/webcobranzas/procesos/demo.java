/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.procesos;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.Pipeline;
import com.itextpdf.tool.xml.XMLWorker;
import com.itextpdf.tool.xml.XMLWorkerFontProvider;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import com.itextpdf.tool.xml.css.CssFilesImpl;
import com.itextpdf.tool.xml.css.StyleAttrCSSResolver;
import com.itextpdf.tool.xml.html.CssAppliersImpl;
import com.itextpdf.tool.xml.html.TagProcessorFactory;
import com.itextpdf.tool.xml.html.Tags;
import com.itextpdf.tool.xml.parser.XMLParser;
import com.itextpdf.tool.xml.pipeline.css.CssResolverPipeline;
import com.itextpdf.tool.xml.pipeline.end.PdfWriterPipeline;
import com.itextpdf.tool.xml.pipeline.html.HtmlPipeline;
import com.itextpdf.tool.xml.pipeline.html.HtmlPipelineContext;
import com.itextpdf.tool.xml.html.HTML;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.Charset;
import pop.webcobranzas.util.ImageTagProcessor;

/**
 *
 * @author Jyoverar
 */
public class demo {

    public static final String DEST = "C:\\pop\\webcobranzas\\resources\\template\\html_12.pdf";
    public static final String HTMLb = "C:\\pop\\webcobranzas\\resources\\template\\templateCarta2.html";

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception, DocumentException {
        // TODO code application logic here
        System.out.println("pop.webcobranzas.procesos.demo.main()");

        File file = new File(DEST);
        file.getParentFile().mkdirs();
        new demo().createPdf(DEST);
    }

    public void createPdf(String files) throws IOException, DocumentException {

        try {
            final OutputStream file = new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\html_12.pdf"));
            final Document document = new Document();
            final PdfWriter writer = PdfWriter.getInstance(document, file);
            document.open();
            final TagProcessorFactory tagProcessorFactory = Tags.getHtmlTagProcessorFactory();
            tagProcessorFactory.removeProcessor(HTML.Tag.IMG);
            tagProcessorFactory.addProcessor(new ImageTagProcessor(), HTML.Tag.IMG);

            final CssFilesImpl cssFiles = new CssFilesImpl();
            cssFiles.add(XMLWorkerHelper.getInstance().getDefaultCSS());
            final StyleAttrCSSResolver cssResolver = new StyleAttrCSSResolver(cssFiles);
            final HtmlPipelineContext hpc = new HtmlPipelineContext(new CssAppliersImpl(new XMLWorkerFontProvider()));
            hpc.setAcceptUnknown(true).autoBookmark(true).setTagFactory(tagProcessorFactory);
            final HtmlPipeline htmlPipeline = new HtmlPipeline(hpc, new PdfWriterPipeline(document, writer));
            final Pipeline<?> pipeline = new CssResolverPipeline(cssResolver, htmlPipeline);
            final XMLWorker worker = new XMLWorker(pipeline, true);
            final Charset charset = Charset.forName("UTF-8");
            final XMLParser xmlParser = new XMLParser(true, worker, charset);
            final InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\templateCarta1.html");
            xmlParser.parse(is, charset);

            is.close();
            document.close();
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
            // TODO
        }

//        final OutputStream file = new FileOutputStream(new File("C:\\pop\\webcobranzas\\resources\\template\\html_12.pdf"));
//
//        // step 1
//        Document document = new Document();
//        // step 2
//        //PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(file));
//        final PdfWriter writer = PdfWriter.getInstance(document, file);
//
//        document.open();
//
//        final TagProcessorFactory tagProcessorFactory = Tags.getHtmlTagProcessorFactory();
//        tagProcessorFactory.removeProcessor(HTML.Tag.IMG);
//        tagProcessorFactory.addProcessor(new ImageTagProcessor(), HTML.Tag.IMG);
//
//        final CssFilesImpl cssFiles = new CssFilesImpl();
//        cssFiles.add(XMLWorkerHelper.getInstance().getDefaultCSS());
//        final StyleAttrCSSResolver cssResolver = new StyleAttrCSSResolver(cssFiles);
//        final HtmlPipelineContext hpc = new HtmlPipelineContext(new CssAppliersImpl(new XMLWorkerFontProvider()));
//        hpc.setAcceptUnknown(true).autoBookmark(true).setTagFactory(tagProcessorFactory);
//        final HtmlPipeline htmlPipeline = new HtmlPipeline(hpc, new PdfWriterPipeline(document, writer));
//        final Pipeline<?> pipeline = new CssResolverPipeline(cssResolver, htmlPipeline);
//        final XMLWorker worker = new XMLWorker(pipeline, true);
//        final Charset charset = Charset.forName("ISO-8859-1");
//        final XMLParser xmlParser = new XMLParser(true, worker, charset);
//        final InputStream is = new FileInputStream("C:\\pop\\webcobranzas\\resources\\template\\templateCarta1.html");
//        xmlParser.parse(is, charset);
//
//        is.close();
//        document.close();
//        file.close();

    }
}
