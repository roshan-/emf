package com.xauto.ux;

import org.apache.commons.lang.StringUtils;

import com.aspose.words.Cell;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.DrawingML;
import com.aspose.words.IImageData;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Row;
import com.aspose.words.Run;
import com.aspose.words.Shape;
import com.aspose.words.ShapeRenderer;
import com.xauto.ux.Emf2.AsposeDrawingType;

public class WordDrawing extends Node{
	final Node node;
	final AsposeDrawingType type;
	
	Node getNode() {
		return node;
	}
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
		String drawingName= node instanceof DrawingML?((DrawingML)node).getName():
							node instanceof Shape?((Shape)node).getName():null;
		if (StringUtils.isBlank(drawingName)) {
			if (node instanceof Shape) {
				Shape s= (Shape)node;
				Node cell= null;
				Node parent= s.getParentNode();
				while(parent != null && parent.getNodeType()!=NodeType.ROW) {
					if (parent.getNodeType()==NodeType.CELL) {
						cell= parent;
					}
					parent= parent.getParentNode();
				}
				if (parent == null) {
					return drawingName;
				}
				Row picturesRow= (Row)parent;
				
				//Figure out which cell index the picture is present
				Node[] currentPictureRowCells= picturesRow.getChildNodes(NodeType.CELL, true).toArray();
				if (currentPictureRowCells == null) {
					//This shouldn't happen, but we have to take care of the possibility
					return drawingName;
				}
				int currentPictureCellIndex= 0;
				for(Node n : currentPictureRowCells) {
					if(n==cell) {
						break;
					}
					currentPictureCellIndex++;
				}
				
				//Assume captions row is the previous row
				Row captionsRow= (Row)picturesRow.getPreviousSibling();
				if (captionsRow == null) {
					return drawingName;
				}
				Cell captionCell= (Cell)captionsRow.getChild(NodeType.CELL, currentPictureCellIndex, true);
				if (captionCell == null) {
					// We didn't find the corresponding caption Cell in the previous Row
					return drawingName;
				}
				
				// Get the Paragraph and all the Text
				StringBuilder sb= new StringBuilder();
				Paragraph[] ps= captionCell.getParagraphs().toArray();
				for (Paragraph p : ps) {
					Run[] rs= p.getRuns().toArray();
					for (Run r : rs) {
						r.getDirectRunAttrsCount();
						sb.append(r.getText());
					}
				}
				// Ignore String Instrumentation
				drawingName= sb.toString().replace("SEQ Figure \\* ARABIC ", "");
			}
		}
		return drawingName;
		
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
