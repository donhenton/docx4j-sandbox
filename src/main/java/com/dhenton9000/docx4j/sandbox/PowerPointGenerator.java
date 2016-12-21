package com.dhenton9000.docx4j.sandbox;

import com.dhenton9000.xml.utils.OffsetAdjuster;
import com.dhenton9000.xml.utils.XmlUtilities;
import java.awt.Dimension;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import org.apache.commons.io.IOUtils;
import org.docx4j.openpackaging.packages.Filetype;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.Pic;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 *
 * @author dhenton
 */
public class PowerPointGenerator {

    private static Logger LOG = LoggerFactory.getLogger(PowerPointGenerator.class);
    private static final String PPTX_TEMPLATE = "/sample-docs/substitution_sample.pptx";
    private static final String PIC_TEMPLATE = "templates/picElement.xml";

    public void generate(HashMap<String, String> mappings, InputStream isImage, OutputStream outStream, String imageSuffix) throws Exception {

        InputStream is = this.getClass().getResourceAsStream(PPTX_TEMPLATE);

        isImage.mark(0);
        Dimension imgDim = getImgDimension(isImage,imageSuffix);
        isImage.reset();
        LOG.debug("dim " + imgDim);
        if (is == null) {
            throw new RuntimeException("can't find template file " + PPTX_TEMPLATE);
        }
        PresentationMLPackage presentationMLPackage
                = (PresentationMLPackage) OpcPackage.load(is, Filetype.ZippedPackage);
        SlidePart slidePart0 = presentationMLPackage.getMainPresentationPart().getSlide(0);
        SlidePart slidePart1 = presentationMLPackage.getMainPresentationPart().getSlide(1);
        List<Object> contents = slidePart1.getContents().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame();
        Iterator partIter = contents.iterator();

        //
        slidePart0.variableReplace(mappings);
        while (partIter.hasNext()) {

            Object i = partIter.next();
            // LOG.debug(i.getClass().getName());
            if (i instanceof Pic) {
                partIter.remove();
            }
        }

        byte[] bytes = IOUtils.toByteArray(isImage);
        BinaryPartAbstractImage newImage
                = BinaryPartAbstractImage.createImagePart(presentationMLPackage, slidePart1, bytes);

        contents.add(1, createPicture(newImage.getSourceRelationships().get(0).getId(),imgDim));
        //String xml = slidePart1.getXML();
        //LOG.debug(xml);
        slidePart1.variableReplace(mappings);

        presentationMLPackage.save(outStream);

    }

    private static Dimension getImgDimension(InputStream is,String suffix) throws IOException {
        ImageInputStream stream = ImageIO.createImageInputStream(is);
        

        Iterator<ImageReader> iter = ImageIO.getImageReadersBySuffix(suffix);

        while (iter.hasNext()) {
            ImageReader reader = iter.next();
            try {

                reader.setInput(stream);
                int width = reader.getWidth(reader.getMinIndex());
                int height = reader.getHeight(reader.getMinIndex());
                return new Dimension(width, height);

            } finally {
                reader.dispose();
            }
        }

        return null;
    }

    private Object createPicture(String relId,Dimension imgDim) throws Exception {
        java.util.HashMap<String, String> mappings = new java.util.HashMap<String, String>();

        OffsetAdjuster oA = new OffsetAdjuster(imgDim,25.0f);
        LOG.debug("oa "+oA.toString());
        
        mappings.put("id1", "4");
        mappings.put("name", "Picture 3");
        mappings.put("descr", "embedded.png");
        mappings.put("rEmbedId", relId);
        mappings.put("offx", oA.getOffsetX());
        mappings.put("offy", oA.getOffsetY());
        mappings.put("extcx", oA.getExtcX());//50% is 5000000
        mappings.put("extcy", oA.getExtcY());


        String pixTemplate
                = XmlUtilities.getStringResource(PIC_TEMPLATE, this.getClass().getClassLoader());
        return org.docx4j.XmlUtils.unmarshallFromTemplate(pixTemplate,
                mappings, Context.jcPML, Pic.class);
    }

    public static void main(String[] args) {
        HashMap<String, String> mappings = new HashMap<String, String>();
        mappings.put("MAIN_TITLE", "DON'T GET A JOB!!!!!");
        mappings.put("SUB_TITLE", "Hang out at Bob's!!!!!");
        mappings.put("IMAGE_TITLE", "Meet the New Boss");
        mappings.put("IMAGE_TEXT", "Same as the Old Boss");
        try {
            InputStream isImage = PowerPointGenerator.class.getResourceAsStream("/sample-docs/p1.jpg");
            // File f = new File();
            FileOutputStream fOut = new FileOutputStream(System.getProperty("user.dir") + "/docs/out/sub200.pptx", false);
            PowerPointGenerator gen = new PowerPointGenerator();

            gen.generate(mappings, isImage, fOut,"jpg");
        } catch (Exception ex) {
            LOG.error("General error in main", ex);
        }

    }
}