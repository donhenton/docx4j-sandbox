/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dhenton9000.docx4j.pptx.sandbox;

import com.dhenton9000.xml.utils.XmlUtilities;
import java.io.File;
import java.io.InputStream;
import org.pptx4j.pml.Shape;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import javax.xml.bind.JAXBException;
import org.apache.commons.io.IOUtils;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.contenttype.ContentTypeManager;
import org.docx4j.openpackaging.contenttype.ContentTypes;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.Filetype;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideLayoutPart;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.pptx4j.Pptx4jException;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.Pic;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

/**
 *
 * @author dhenton
 */
public class PPTXExplorer {

    protected static Logger LOG = LoggerFactory.getLogger(PPTXExplorer.class);

    private static boolean MACRO_ENABLE = false;

    public static void main(String[] args) throws Exception {

        //simpleCreate();
        PPTXExplorer p = new PPTXExplorer();
       p.replaceImage();
      // p.getXML();

    }
    
    private void getXML() throws Exception
    {
        
        
        String SAMPLE_PICTURE =  
        XmlUtilities.getStringResource("templates/picElement.xml",this.getClass().getClassLoader());
        //Document doc = XmlUtilities.stringToDoc(SAMPLE_PICTURE);
        //String ii = XmlUtilities.docToString(doc, true,true);
        LOG.debug("\n\n"+SAMPLE_PICTURE+"\n\n");
        
    }

    private void replaceImage() throws Exception {

        InputStream is = this.getClass().getResourceAsStream("/sample-docs/substitution_sample.pptx");
        if (is == null) {
            throw new RuntimeException("can't find file");
        }
        PresentationMLPackage presentationMLPackage
                = (PresentationMLPackage) OpcPackage.load(is, Filetype.ZippedPackage);
        SlidePart slidePart = presentationMLPackage.getMainPresentationPart().getSlide(1);
        List<Object> contents = slidePart.getContents().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame();
        Iterator partIter = contents.iterator();

        while (partIter.hasNext()) {

            Object i = partIter.next();
           // LOG.debug(i.getClass().getName());
            if (i instanceof Pic) {
                partIter.remove();
            }
        }
        
        InputStream isImage = this.getClass().getResourceAsStream("/sample-docs/p1.jpg");
        byte[] bytes = IOUtils.toByteArray(isImage);
        BinaryPartAbstractImage newImage
                = BinaryPartAbstractImage.createImagePart(presentationMLPackage, slidePart, bytes);

        contents.add(1, createPicture(newImage.getSourceRelationships().get(0).getId()));
        //String xml = slidePart.getXML();
        //LOG.debug(xml);

        File f = new File(System.getProperty("user.dir") + "/docs/out/sub.pptx");
        presentationMLPackage.save(f);

    }

    private void replaceText() throws Docx4JException, Pptx4jException, JAXBException {

        InputStream is = this.getClass().getResourceAsStream("/sample-docs/substitution_sample.pptx");
        if (is == null) {
            throw new RuntimeException("can't find file");
        }
        PresentationMLPackage presentationMLPackage
                = (PresentationMLPackage) OpcPackage.load(is, Filetype.ZippedPackage);
        SlidePart slidePart = presentationMLPackage.getMainPresentationPart().getSlide(0);
        HashMap<String, String> mappings = new HashMap<String, String>();
        //${MAIN_TITLE}, ${SUB_TITLE} are in the file as markers

        mappings.put("MAIN_TITLE", "DON'T GET A JOB!!!!!");
        mappings.put("SUB_TITLE", "Hang out at Bob's!!!!!");
        slidePart.variableReplace(mappings);

        //String xml = slidePart.getXML();
        //LOG.debug(xml);
        File f = new File(System.getProperty("user.dir") + "/docs/out/sub.pptx");
        presentationMLPackage.save(f);

    }

    private void simpleCreate() throws Pptx4jException, Docx4JException, InvalidFormatException, JAXBException, URISyntaxException {
        // Where will we save our new .ppxt?
        String outputfilepath = System.getProperty("user.dir") + "/docs/out/pptx-test.pptx";
        if (MACRO_ENABLE) {
            outputfilepath += "m";
        }

        // Create skeletal package, including a MainPresentationPart and a SlideLayoutPart
        PresentationMLPackage presentationMLPackage = PresentationMLPackage.createPackage();

        if (MACRO_ENABLE) {
            ContentTypeManager ctm = presentationMLPackage.getContentTypeManager();
            ctm.removeContentType(new PartName("/ppt/presentation.xml"));
            ctm.addOverrideContentType(new URI("/ppt/presentation.xml"), ContentTypes.PRESENTATIONML_MACROENABLED);
        }

        // Need references to these parts to create a slide
        // Please note that these parts *already exist* - they are
        // created by createPackage() above.  See that method
        // for instruction on how to create and add a part.
        MainPresentationPart pp = (MainPresentationPart) presentationMLPackage.getParts().getParts().get(
                new PartName("/ppt/presentation.xml"));
        SlideLayoutPart layoutPart = (SlideLayoutPart) presentationMLPackage.getParts().getParts().get(
                new PartName("/ppt/slideLayouts/slideLayout1.xml"));

        // OK, now we can create a slide
        SlidePart slidePart = new SlidePart(new PartName("/ppt/slides/slide1.xml"));
        slidePart.setContents(SlidePart.createSld());
        pp.addSlide(0, slidePart);

        // Slide layout part
        slidePart.addTargetPart(layoutPart);

        // String sampleShape = null;
        String sampleShape
                = "<p:sp   xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
                + "<p:nvSpPr>"
                + "<p:cNvPr id=\"4\" name=\"Title 3\" />"
                + "<p:cNvSpPr>"
                + "<a:spLocks noGrp=\"1\" />"
                + "</p:cNvSpPr>"
                + "<p:nvPr>"
                + "<p:ph type=\"title\" />"
                + "</p:nvPr>"
                + "</p:nvSpPr>"
                + "<p:spPr />"
                + "<p:txBody>"
                + "<a:bodyPr />"
                + "<a:lstStyle />"
                + "<a:p>"
                + "<a:r>"
                + "<a:rPr lang=\"en-US\" smtClean=\"0\" />"
                + "<a:t>Get a Job, you worthless Slob!</a:t>"
                + "</a:r>"
                + "<a:endParaRPr lang=\"en-US\" />"
                + "</a:p>"
                + "</p:txBody>"
                + "</p:sp>";
        // Create and add shape
        Shape sample = ((Shape) XmlUtils.unmarshalString(sampleShape, Context.jcPML));
        slidePart.getContents().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(sample);

        // All done: save it
        presentationMLPackage.save(new java.io.File(outputfilepath));

        System.out.println("\n\n done .. saved " + outputfilepath);
    }
    
    
    
    private Object createPicture(String relId) throws Exception
    {
          java.util.HashMap<String, String>mappings = new java.util.HashMap<String, String>();
        
        mappings.put("id1", "4");
        mappings.put("name", "Picture 3");
        mappings.put("descr", "embedded.png");
        mappings.put("rEmbedId", relId );
        mappings.put("offx", Long.toString(0));
        mappings.put("offy", Long.toString(0));
        mappings.put("extcx", Long.toString(7143750));
        mappings.put("extcy", Long.toString(7143750));
         String pixTemplate =  
        XmlUtilities.getStringResource("templates/picElement.xml",this.getClass().getClassLoader());
        return org.docx4j.XmlUtils.unmarshallFromTemplate(pixTemplate, 
        		mappings, Context.jcPML, Pic.class ) ;
    }
    
    
}
