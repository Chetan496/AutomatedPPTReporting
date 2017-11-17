package com.testpptx4j;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.UnsupportedEncodingException;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.dml.CTPoint2D;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.CTTable;
import org.docx4j.dml.CTTableCell;
import org.docx4j.dml.CTTableCol;
import org.docx4j.dml.CTTableGrid;
import org.docx4j.dml.CTTableRow;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.GraphicData;
import org.docx4j.dml.ObjectFactory;
import org.docx4j.docProps.core.dc.terms.URI;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideLayoutPart;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.pptx4j.Pptx4jException;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.CTGraphicalObjectFrame;
import org.pptx4j.pml.Sld;
import org.pptx4j.pml.*;

public class LearningPPTX4J {

	private static String SAMPLE_SHAPE = 			
			"<p:sp   xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
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
					+ "<a:t>Hello World</a:t>"
				+ "</a:r>"
				+ "<a:endParaRPr lang=\"en-US\" />"
			+ "</a:p>"
		+ "</p:txBody>"
	+ "</p:sp>";

	
	public static void main(String[] args)
	{
		
		String templateFile = "E:\\DocX4J_data\\template.pptx";
		String outFile = "E:\\DocX4J_data\\output.pptx";
		
		try {
			InputStream inputStream = new FileInputStream(templateFile);
			PresentationMLPackage presentationMLPackage= 
					(PresentationMLPackage) OpcPackage.load(inputStream);
			
			MainPresentationPart pp = (MainPresentationPart)
                    presentationMLPackage.getParts().getParts().get(new
                    PartName("/ppt/presentation.xml"));
			
			SlideLayoutPart layoutPart = (SlideLayoutPart)presentationMLPackage.getParts().getParts().get(
					new PartName("/ppt/slideLayouts/slideLayout7.xml"));
			
			//lets create a new slide using the slidelayout part
			SlidePart slidePart = new SlidePart( new PartName("/ppt/slides/slide" + 2+ ".xml"));
			slidePart.setContents( SlidePart.createSld() ); 
			
			pp.addSlide(slidePart);
			slidePart.addTargetPart(layoutPart);

			List shapeTree = slidePart.getContents().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame();	
			CTGraphicalObjectFrame graphicFrame = CreateTableInPowerPoint(4, 10, 1);
			shapeTree.add(graphicFrame);

			
			 presentationMLPackage.save(new File(outFile));
			 

			 inputStream.close();
			 
			System.out.println("Done");
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Pptx4jException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} /*catch (JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}*/ catch (JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		
		
		
	}
	
	
	private static void SetSlideContentBasedOnXML(SlidePart slidePart) {
		/*Demonstartes how to add content to slide by unmarshalling xml string */
		
		StringBuffer slideXMLBuffer = new StringBuffer();
		BufferedReader br = null;
		String line = "";
		
		String slideDataXmlFile = "E:\\DocX4J_data\\xmldata_for_slide.xml";
		
		InputStream xmlDataInputStream;
		try {
			xmlDataInputStream = new FileInputStream(slideDataXmlFile);
			Reader fr = new InputStreamReader(xmlDataInputStream, "utf-8");
			br = new BufferedReader(fr);
			while ((line = br.readLine()) != null) {
			           slideXMLBuffer.append(line);
			          slideXMLBuffer.append(" ");
			}
			
			Sld sld = (Sld) XmlUtils.unmarshalString(slideXMLBuffer.toString(),Context.jcPML,
		               Sld.class);
			
			slidePart.setJaxbElement(sld);
			xmlDataInputStream.close();
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		
		
	}
	
	
	
	
	/*
	 * 
	 * 
	 *  List shapeTree = slideTemplate.getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame();
      CTGraphicFrame gf = new CTGraphicFrame();
      shapeTree.add(gf);
	 */
	
	private static CTGraphicalObjectFrame CreateTableInPowerPoint(int rows, int cols, int cellWidthTips)
	{
	
		org.pptx4j.pml.ObjectFactory graphicObjectFactory = Context.getpmlObjectFactory();
		org.docx4j.dml.ObjectFactory dmlObjectFactory = new ObjectFactory();
		
		CTGraphicalObjectFrame graphicFrame = graphicObjectFactory.createCTGraphicalObjectFrame();
		org.pptx4j.pml.CTGraphicalObjectFrameNonVisual nvGraphicFramePr=graphicObjectFactory.createCTGraphicalObjectFrameNonVisual();
	    org.docx4j.dml.CTNonVisualDrawingProps cNvPr=dmlObjectFactory.createCTNonVisualDrawingProps();
	    org.docx4j.dml.CTNonVisualGraphicFrameProperties cNvGraphicFramePr=dmlObjectFactory.createCTNonVisualGraphicFrameProperties();
	    org.docx4j.dml.CTGraphicalObjectFrameLocking graphicFrameLocks=new org.docx4j.dml.CTGraphicalObjectFrameLocking();
	    org.docx4j.dml.CTTransform2D xfrm=dmlObjectFactory.createCTTransform2D();
		
		
	    /*TODO: check how styling can be given to the table*/
	
		Graphic graphic = dmlObjectFactory.createGraphic();		
		GraphicData graphicData = dmlObjectFactory.createGraphicData();
		graphicData.setUri("http://schemas.openxmlformats.org/drawingml/2006/table");

	
		
		CTTable ctTable = dmlObjectFactory.createCTTable() ;
		JAXBElement<CTTable> tbl = dmlObjectFactory.createTbl(ctTable);	
		CTTableGrid ctTableGrid = dmlObjectFactory.createCTTableGrid();
	    CTTableCol gridCol = dmlObjectFactory.createCTTableCol();
	    CTTableRow ctTableRow = dmlObjectFactory.createCTTableRow();
	    
	    //Build the parent-child relationship of this slides.xml
	    graphicFrame.setNvGraphicFramePr(nvGraphicFramePr);
	    nvGraphicFramePr.setCNvPr(cNvPr);
	    cNvPr.setName("Table 1");
	    cNvPr.setId(3L);
	    nvGraphicFramePr.setCNvGraphicFramePr(cNvGraphicFramePr);
        cNvGraphicFramePr.setGraphicFrameLocks(graphicFrameLocks);
        graphicFrameLocks.setNoGrp(true);
        nvGraphicFramePr.setNvPr(graphicObjectFactory.createNvPr());
        graphicFrame.setXfrm(xfrm);
        
        CTPositiveSize2D ext = dmlObjectFactory.createCTPositiveSize2D();
        ext.setCx(6096000);
        ext.setCy(741680);
        xfrm.setExt(ext);
        
        CTPoint2D off = dmlObjectFactory.createCTPoint2D(); //offset
        xfrm.setOff(off);
        off.setX(1524000);
        off.setY(1397000);
        
        graphicFrame.setGraphic(graphic);
        graphic.setGraphicData(graphicData);
        graphicData.setUri("http://schemas.openxmlformats.org/drawingml/2006/table");
        graphicData.getAny().add(tbl);
        
        ctTable.setTblGrid(ctTableGrid);
        for(int i=0; i < cols; i++)
    	{
        	   ctTableGrid.getGridCol().add(gridCol);	
    	}
        gridCol.setW(600000);
        
        ctTableRow.setH(370840);
        try {

        	for(int i=0; i < cols; i++)
         	{
        			ctTableRow.getTc().add(createTableCell());	
         	}

		} catch (JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
        //we are repeating the rows. each row has N cells
        for(int i=0;i < rows;i++){
        	ctTable.getTr().add(ctTableRow);
        }
     
	    return graphicFrame;
	     
		
	}
	
	
	public static CTTableCell createTableCell() throws JAXBException {
		   String contents =
			"<a:tc  xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
			"<a:txBody>"
		        +"<a:bodyPr/>"
		        +"<a:lstStyle/>"
		        +"<a:p>"
		          +"<a:r>"
		            +"<a:rPr lang=\"en-AU\" dirty=\"0\" smtClean=\"0\"/>"
		            +"<a:t>12</a:t>"
		          +"</a:r>"
		          +"<a:endParaRPr lang=\"en-AU\" dirty=\"0\"/>"
		          +"</a:p>"
		      +"</a:txBody>" +
		      "</a:tc>";
		      //+"<a:tcPr/>
		   return ((CTTableCell)XmlUtils.unmarshalString(contents,org.docx4j.jaxb.Context.jc, CTTableCell.class));
			
		}
}
