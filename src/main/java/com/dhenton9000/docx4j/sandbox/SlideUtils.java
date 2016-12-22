/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dhenton9000.docx4j.sandbox;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideLayoutPart;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.Sld;

/**
 * PresentationMLPackage presentationMLPackage = (PresentationMLPackage)
 * OpcPackage.load(is, Filetype.ZippedPackage);
 *
 * @author dhenton
 */
public class SlideUtils {

    public static final String MAIN_PRESENTATION_NAME = "/ppt/presentation.xml";
    private static final String SLIDE_TEMPLATE = "/part_templates/slide_part.xml";
    
    
    /**
     * This will insert a blank slide using SLIDE_TEMPLATE as a template.
     * The template contains variables suitable for replacement
     * 
     * TODO
     * add the stuff needed to variable replacement
     * 
     * @param presentationMLPackage the pointer to the original pptx file
     * @param slideIndex the index of where you want the slide to be added
     * so if three slides already there, slideIndex should be 3  (starts a 0)
     * @throws Exception 
     */
    public static void insertSlide(PresentationMLPackage presentationMLPackage, int slideIndex)
            throws Exception {

        MainPresentationPart mainPresentationPart = (MainPresentationPart) presentationMLPackage
                .getParts().getParts().get(new PartName(MAIN_PRESENTATION_NAME));

        SlideLayoutPart layoutPart = (SlideLayoutPart) presentationMLPackage.getParts().getParts()
                .get(new PartName("/ppt/slideLayouts/slideLayout2.xml"));

         
        SlidePart slide1 =  PresentationMLPackage.createSlidePart(mainPresentationPart, layoutPart,
                new PartName("/ppt/slides/slide" + slideIndex + ".xml"));
      //  SlidePart slide1 = new SlidePart(new PartName("/ppt/slides/slide" + slideIndex + ".xml"));
 //       slide1.addTargetPart(layoutPart);
       
       //  mainPresentationPart.addSlide(slideIndex,slide1);
         
        
        StringBuilder slideXMLBuffer = new StringBuilder();
        BufferedReader br = null;
        String line = "";

        InputStream in = SlideUtils.class.getResourceAsStream(SLIDE_TEMPLATE);
        Reader fr = new InputStreamReader(in, "utf-8");
        br = new BufferedReader(fr);
        while ((line = br.readLine()) != null) {
            slideXMLBuffer.append(line);
            slideXMLBuffer.append(" ");
        }
        Sld sld = (Sld) XmlUtils.unmarshalString(slideXMLBuffer.toString(), Context.jcPML,
                Sld.class);
        //slide1.setJaxbElement(sld);
        slide1.setContents(sld);
        
    }

}
