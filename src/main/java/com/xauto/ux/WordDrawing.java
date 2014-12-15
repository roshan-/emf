package com.xauto.ux;

import com.aspose.words.DocumentVisitor;
import com.aspose.words.DrawingML;
import com.aspose.words.IImageData;
import com.aspose.words.Node;
import com.aspose.words.Shape;
import com.aspose.words.ShapeRenderer;
import com.xauto.ux.Emf2.AsposeDrawingType;

public class WordDrawing extends Node{
	final Node node;
	final AsposeDrawingType type;
	
	WordDrawing(DrawingML node) {
		this.node= node;
		type= AsposeDrawingType.DRAWING_ML;
	}
	
	WordDrawing(Shape node) {
		this.node= node;
		type= AsposeDrawingType.SHAPE;
	}

	@Override
	public boolean accept(DocumentVisitor arg0) throws Exception {
		if (node instanceof DrawingML) {
			return ((DrawingML)node).accept(arg0);
		} else if (node instanceof Shape) {
			return ((Shape)node).accept(arg0);
		} else {
			return false;
		}
	}

	@Override
	public int getNodeType() {
		if (node instanceof DrawingML) {
			return ((DrawingML)node).getNodeType();
		} else if (node instanceof Shape) {
			return ((Shape)node).getNodeType();
		} else {
			return 0;
		}
	}
	
	@Override
	public String toString() {
		if (node instanceof DrawingML) {
			return ((DrawingML)node).toString();
		} else if (node instanceof Shape) {
			return ((Shape)node).toString();
		} else {
			return null;
		}
	}
	
	public String getName() {
		if (node instanceof DrawingML) {
			return ((DrawingML)node).getName();
		} else if (node instanceof Shape) {
			return ((Shape)node).getName();
		} else {
			return null;
		}
	}

	public ShapeRenderer getShapeRenderer() throws Exception {
		if (node instanceof DrawingML) {
			return ((DrawingML)node).getShapeRenderer();
		} else if (node instanceof Shape) {
			return ((Shape)node).getShapeRenderer();
		} else {
			return null;
		}
	}
	
	public IImageData getImageData() {
		if (node instanceof DrawingML) {
			return ((DrawingML)node).getImageData();
		} else if (node instanceof Shape) {
			return ((Shape)node).getImageData();
		} else {
			return null;
		}
	}

}
