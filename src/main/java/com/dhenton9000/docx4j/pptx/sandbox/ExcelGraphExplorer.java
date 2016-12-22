/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dhenton9000.docx4j.pptx.sandbox;

import com.dhenton9000.docx4j.sandbox.SlideUtils;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;
import javax.xml.bind.JAXBException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.Filetype;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.utils.BufferUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xlsx4j.sml.Row;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;

/**
 * http://bridgei2i.com/blog/programmatically-creating-ms-office-compatible-charts/
 * https://github.com/plutext/docx4j/blob/master/src/samples/pptx4j/org/pptx4j/samples/EditEmbeddedCharts.java
 * http://stackoverflow.com/questions/30556157/printing-contents-of-xlsx-sheet
 * (use of shared string for xlsx text cells)
 *
 * @author dhenton
 */
public class ExcelGraphExplorer {

    protected static Logger LOG = LoggerFactory.getLogger(PPTXExplorer.class);

    private static boolean MACRO_ENABLE = false;

    public static void main(String[] args) throws Exception {

        //simpleCreate();
        ExcelGraphExplorer p = new ExcelGraphExplorer();
        p.addSlide();

    }

    public void addSlide() throws Exception {

        InputStream is = this.getClass().getResourceAsStream("/sample-docs/graph_sample.pptx");
        if (is == null) {
            throw new RuntimeException("can't find file");
        }
        PresentationMLPackage presentationMLPackage
                = (PresentationMLPackage) OpcPackage.load(is, Filetype.ZippedPackage);
        HashMap<String, String> mappings = new HashMap<String, String>();
        mappings.put("GRAPH_TITLE", "DON'T GET A JOB!!!!!");
        mappings.put("MAIN_TEXT", "Hang out at Bob's!!!!!");
        SlideUtils.appendSlide(presentationMLPackage, mappings);
        File f = new File(System.getProperty("user.dir") + "/docs/out/graph_demo.pptx");
        presentationMLPackage.save(f);

    }

    public void createGraph() throws Docx4JException, IOException, InvalidFormatException, JAXBException {
        String templateFile = "";

        InputStream is = this.getClass().getResourceAsStream("/sample-docs/graph_with_chart.pptx");
        if (is == null) {
            throw new RuntimeException("can't find file");
        }
        PresentationMLPackage presentationMLPackage
                = (PresentationMLPackage) OpcPackage.load(is, Filetype.ZippedPackage);

        MainPresentationPart pp = (MainPresentationPart) presentationMLPackage
                .getParts().getParts().get(new PartName("/ppt/presentation.xml"));

        EmbeddedPackagePart spreadsheetPart
                = (EmbeddedPackagePart) presentationMLPackage.getParts().getParts()
                        .get(new PartName("/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));

        InputStream isXls = BufferUtil.newInputStream(spreadsheetPart.getBuffer());

        SpreadsheetMLPackage spreadSheet = (SpreadsheetMLPackage) SpreadsheetMLPackage.load(isXls);

        spreadSheet.getParts()
                .getParts().forEach((key, value) -> {

                    //  LOG.debug(String.format("key %s part Name %s", key,value.getPartName()) );
                });

        WorksheetPart worksheetPart = (WorksheetPart) spreadSheet.getParts().getParts().entrySet().stream().filter(entryPart -> {
            return entryPart.getValue() instanceof WorksheetPart;

        }).map(e -> {
            return e.getValue();
        }).collect(Collectors.toList()).get(0);
        List<Row> originalRows = worksheetPart.getContents().getSheetData().getRow();

        SharedStrings sharedStrings = (SharedStrings) spreadSheet.getParts().getParts().entrySet().stream().filter(entryPart -> {
            return entryPart.getValue() instanceof SharedStrings;

        }).map(e -> {
            return e.getValue();
        }).collect(Collectors.toList()).get(0);

        LOG.debug("shared " + sharedStrings.toString());

        List<Row> rows = originalRows.stream().filter(row -> {
            return row.getC().size() > 0;
        }).map(row -> {

            // LOG.debug(row.getC().size()+"");
            return row;
        }).collect(Collectors.toList());

        rows.forEach(row -> {
            row.getC().forEach(cell -> {

                LOG.debug(String.format("row %s col %s value %s  type %s", row.getR() + "", cell.getR(), cell.getV(), cell.getT().toString()));

            });

        });

        //LOG.debug(spreadsheetPart.getPackage().getParts().getPart());
//        spreadsheetPart.getPackage().getParts().getParts().forEach((key,  value)->{
//            
//           LOG.debug(String.format("key %s part Name %s", key,value.getPartName()) );
//
//            
//       }) ;
//        SlideLayoutPart layoutPart = (SlideLayoutPart) presentationMLPackage.getParts().getParts()
//                .get(new PartName("/ppt/slideLayouts/slideLayout2.xml"));
//        String slideIndex = "2";
//
//        SlidePart slide1 = PresentationMLPackage.createSlidePart(pp, layoutPart,
//                new PartName("/ppt/slides/slide" + slideIndex + ".xml"));
//        StringBuilder slideXMLBuffer = new StringBuilder();
//        BufferedReader br = null;
//        String line = "";
//        String slideDataXmlFile = "/part_templates/slide_part.xml";
//        InputStream in = this.getClass().getResourceAsStream(slideDataXmlFile);
//        Reader fr = new InputStreamReader(in, "utf-8");
//        br = new BufferedReader(fr);
//        while ((line = br.readLine()) != null) {
//            slideXMLBuffer.append(line);
//            slideXMLBuffer.append(" ");
//        }
//        Sld sld = (Sld) XmlUtils.unmarshalString(slideXMLBuffer.toString(), Context.jcPML,
//                Sld.class);
//        slide1.setJaxbElement(sld);
        ////////////////do the save
        //File f = new File(System.getProperty("user.dir") + "/docs/out/graph_demo.pptx");
        //presentationMLPackage.save(f);
    }

}
