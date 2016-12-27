/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dhenton9000.batik.sandbox;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.batik.dom.svg.SVGDOMImplementation;
import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.batik.transcoder.image.JPEGTranscoder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 *
 * @author dhenton
 */
public class BatikExplorer {

    protected static Logger LOG = LoggerFactory.getLogger(BatikExplorer.class);

    public static void main(String[] args) throws Exception {

        //simpleCreate();
        BatikExplorer p = new BatikExplorer();
        try {
        p.exploreBatik();
        }
        catch(Exception e)
        {
            LOG.error("ERROR: "+e.getMessage(),e);
        }
    }

    private void exploreBatik() throws Exception {
        Document document = createDocument();
        save(document);
        System.exit(0);
    }

    public Document createDocument() {

        // Create a new document.
        DOMImplementation impl = SVGDOMImplementation.getDOMImplementation();
        String svgNS = SVGDOMImplementation.SVG_NAMESPACE_URI;
        Document document
                = impl.createDocument(svgNS, "svg", null);
        Element root = document.getDocumentElement();
        root.setAttributeNS(null, "width", "450");
        root.setAttributeNS(null, "height", "500");

        // Add some content to the document.
        Element e;
        e = document.createElementNS(svgNS, "rect");
        e.setAttributeNS(null, "x", "10");
        e.setAttributeNS(null, "y", "10");
        e.setAttributeNS(null, "width", "200");
        e.setAttributeNS(null, "height", "300");
        e.setAttributeNS(null, "style", "fill:red;stroke:black;stroke-width:4");
        root.appendChild(e);

        e = document.createElementNS(svgNS, "circle");
        e.setAttributeNS(null, "cx", "225");
        e.setAttributeNS(null, "cy", "250");
        e.setAttributeNS(null, "r", "100");
        e.setAttributeNS(null, "style", "fill:green;fill-opacity:.5");
        root.appendChild(e);

        return document;
    }
 
    public void save(Document document) throws Exception {

        // Create a JPEGTranscoder and set its quality hint.
       JPEGTranscoder t = new JPEGTranscoder();
         t.addTranscodingHint(JPEGTranscoder.KEY_QUALITY,
                 new Float(.8));

        // Set the transcoder input and output.
        TranscoderInput input = new TranscoderInput(document);
        try (OutputStream ostream = new FileOutputStream("out.jpg")) {
            TranscoderOutput output = new TranscoderOutput(ostream);
            
            // Perform the transcoding.
            t.transcode(input, output);
            ostream.flush();
        }
        
    }

// 
//    
//      public byte[] svgToPNG(String svg) throws TranscoderException {
//        TranscoderInput transcoderInput = new TranscoderInput(new StringReader(svg));
//        ByteArrayOutputStream output = new ByteArrayOutputStream();
//        TranscoderOutput transcoderOutput = new TranscoderOutput(output);
//        PNGTranscoder pngTranscoder = new PNGTranscoder();
//       
//            pngTranscoder.transcode(transcoderInput, transcoderOutput);
//         
//
//
//        return output.toByteArray();
//    }




}
