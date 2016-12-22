/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.dhenton9000.docx4j.pptx.sandbox;

import com.dhenton9000.docx4j.sandbox.SlideUtils;
import java.io.File;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;
import org.docx4j.dml.chart.CTNumData;
import org.docx4j.dml.chart.CTNumDataSource;
import org.docx4j.dml.chart.CTNumRef;
import org.docx4j.dml.chart.CTNumVal;
import org.docx4j.dml.chart.CTPieChart;
import org.docx4j.dml.chart.CTPieSer;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.Filetype;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.utils.BufferUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xlsx4j.sml.Row;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;
import org.xlsx4j.sml.STCellType;

/**
 * http://bridgei2i.com/blog/programmatically-creating-ms-office-compatible-charts/
 * https://github.com/plutext/docx4j/blob/master/src/samples/pptx4j/org/pptx4j/samples/EditEmbeddedCharts.java
 * http://stackoverflow.com/questions/30556157/printing-contents-of-xlsx-sheet
 * https://github.com/plutext/docx4j/blob/master/src/samples/pptx4j/org/pptx4j/samples/EditEmbeddedCharts.java
 * (use of shared string for xlsx text cells)
 *
 * @author dhenton
 */
public class ExcelGraphExplorer {

    protected static Logger LOG = LoggerFactory.getLogger(ExcelGraphExplorer.class);

    private static boolean MACRO_ENABLE = false;

    public static void main(String[] args) throws Exception {

        //simpleCreate();
        ExcelGraphExplorer p = new ExcelGraphExplorer();
        p.findChartStrRef();

    }

    public void findChartStrRef() throws Exception {

        InputStream is = this.getClass().getResourceAsStream("/sample-docs/graph_with_chart.pptx");
        if (is == null) {
            throw new RuntimeException("can't find file");
        }
        PresentationMLPackage presentationMLPackage
                = (PresentationMLPackage) OpcPackage.load(is, Filetype.ZippedPackage);

        MainPresentationPart pp = (MainPresentationPart) presentationMLPackage
                .getParts().getParts().get(new PartName("/ppt/presentation.xml"));

        Chart chart = (Chart) presentationMLPackage.getParts().get(new PartName("/ppt/charts/chart1.xml"));
      //  LOG.debug(chart.toString());
      //  LOG.debug(chart.getContents().getChart().toString());
        CTPieChart pieChart = (CTPieChart) chart.getContents().getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart().get(0);

        CTNumData pieSeriesValues = pieChart.getSer().get(0).getVal().getNumRef().getNumCache();
        pieSeriesValues.getPt().forEach((CTNumVal d) -> {

            LOG.debug(String.format("index %d value %s", d.getIdx(), d.getV()));

        });

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

    public void readCells() throws Exception {
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

        // LOG.debug("shared " + sharedStrings.toString());
        List<Row> rows = originalRows.stream().filter(row -> {
            return row.getC().size() > 0;
        }).map(row -> {

            // LOG.debug(row.getC().size()+"");
            return row;
        }).collect(Collectors.toList());

        rows.forEach(row -> {
            row.getC().forEach(cell -> {
                try {
                    String output = "not found";
                    if (cell.getT().equals(STCellType.S)) {

                        output = sharedStrings.getContents().getSi().get(Integer.parseInt(cell.getV())).getT().getValue();

                    } else {
                        output = cell.getV();
                    }

                    LOG.debug(String.format("col %s value %s", cell.getR(), output));
                } catch (Docx4JException ex) {
                    throw new RuntimeException(ex);
                }
            });

        });
    }

}
