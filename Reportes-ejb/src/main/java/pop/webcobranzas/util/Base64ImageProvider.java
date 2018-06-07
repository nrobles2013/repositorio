/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pop.webcobranzas.util;

import com.itextpdf.text.BadElementException;
import com.itextpdf.text.Image;
import com.itextpdf.tool.xml.pipeline.html.AbstractImageProvider;
import com.itextpdf.text.pdf.codec.Base64;
import java.io.IOException;

/**
 *
 * @author Jyoverar
 */
public class Base64ImageProvider extends AbstractImageProvider {

    @Override
    public Image retrieve(String src) {
        int pos = src.indexOf("base64,");
        try {
            if (src.startsWith("data") && pos > 0) {
                byte[] img = Base64.decode(src.substring(pos + 7));
                return Image.getInstance(img);
            } else {
                return Image.getInstance(src);
            }
        } catch (BadElementException ex) {
            return null;
        } catch (IOException ex) {
            return null;
        }
    }

    @Override
    public String getImageRootPath() {
        return null;
    }

}
