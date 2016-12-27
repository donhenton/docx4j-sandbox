/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dhenton9000.batik.sandbox;

import com.dhenton9000.docx4j.sandbox.PowerPointGenerator;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import org.apache.batik.dom.svg.SVGDOMImplementation;
import org.apache.commons.io.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.OutputStream;
import java.util.HashMap;
import org.apache.commons.io.FileUtils;

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
            p.testOutBatikTranscoder();
            LOG.info("done");
            System.exit(0);
        } catch (Exception e) {
            LOG.error("MAIN ERROR: " + e.getMessage() + "\n", e);
        }
    }

    private void testOutBatikTranscoder() throws Exception {

        InputStream docStream = this.getClass().getResourceAsStream("/sample-docs/sample.svg");
        String svgInput = IOUtils.toString(docStream);

        D3GraphBatikTransCoder tCoder = new D3GraphBatikTransCoder();
        InputStream isImage = tCoder.loadDocument(svgInput);

        File f = new File(System.getProperty("user.dir") + "/docs/out/batik_out.jpg");
        OutputStream outputStream = new FileOutputStream(f);

        int read = 0;
        byte[] bytes = new byte[1024];

        while ((read = isImage.read(bytes)) != -1) {
            outputStream.write(bytes, 0, read);
        }

    }

    private void testOutPPTX() throws Exception {

        HashMap<String, String> mappings = new HashMap<String, String>();
        mappings.put("MAIN_TITLE", "DON'T GET A JOB!!!!!");
        mappings.put("SUB_TITLE", "Hang out at Bob's!!!!!");
        mappings.put("IMAGE_TITLE", "Meet the New Boss");
        mappings.put("IMAGE_TEXT", "Same as the Old Boss");
        try {
            InputStream docStream = this.getClass().getResourceAsStream("/sample-docs/sample.svg");
            String svgInput = IOUtils.toString(docStream);

            D3GraphBatikTransCoder tCoder = new D3GraphBatikTransCoder();
            InputStream isImage = tCoder.loadDocument(svgInput);

            FileOutputStream fOut = new FileOutputStream(System.getProperty("user.dir") + "/docs/out/svgPresenation.pptx", false);
            PowerPointGenerator gen = new PowerPointGenerator();

            gen.generate(mappings, isImage, fOut, "jpg");
        } catch (Exception ex) {
            LOG.error("General error in main", ex);
        }

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

}
