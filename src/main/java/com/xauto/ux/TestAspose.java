package com.xauto.ux;

import java.io.FileNotFoundException;

import org.freehep.graphicsio.emf.EMF2SVG;

import com.aspose.imaging.Source;
import com.aspose.imaging.exceptions.imageformats.MetafilesException;
import com.aspose.imaging.fileformats.metafile.EmfMetafileImage;
import com.aspose.imaging.imageoptions.PngOptions;
import com.aspose.words.Document;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Shape;
import com.aspose.words.SvgSaveOptions;

public class TestAspose {
	public static void emf2png() throws MetafilesException, FileNotFoundException {
		EmfMetafileImage emf= new EmfMetafileImage("image8.emf");
		PngOptions png= new PngOptions();
	}

	public static void emf2svg() {
		String[] args= new String[1];
		args[0]= "image8.emf";
		EMF2SVG.main(args);
	}
	public static void main(String[] args) throws Exception {
		emf2svg();
		if (true) return;
		System.out.println(System.getProperty("java.class.path"));
		Document doc= new Document("test.docx");
		EmfMetafileImage emf= new EmfMetafileImage("image8.emf");
		PngOptions png= new PngOptions();
		Source value;
		SvgSaveOptions svg= new SvgSaveOptions();
		
		
		emf.save("image81.png", png);
		@SuppressWarnings("unchecked")
		NodeCollection<Node> nodes= doc.getChildNodes();
		for (Node node : nodes) {
			System.out.println(node.getNodeType());
		}
		
		NodeCollection<Shape> shapes= doc.getChildNodes(NodeType.SHAPE, true);
		int index=0;
		for (Shape shape : shapes) {
			String imgFileName= "image"+index+FileFormatUtil.imageTypeToExtension(shape.getImageData().getImageType());
			shape.getImageData().save(imgFileName);
			index++;
		}

	}

}
