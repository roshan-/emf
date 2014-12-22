package com.xauto.ux;

import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.GraphicsEnvironment;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.FilenameFilter;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.imgscalr.Scalr;
import org.imgscalr.Scalr.Method;
import org.imgscalr.Scalr.Mode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.aspose.imaging.exceptions.imageformats.MetafilesException;
import com.aspose.imaging.fileformats.metafile.EmfMetafileImage;
import com.aspose.imaging.imageoptions.PngOptions;
import com.aspose.words.DmlEffectsRenderingMode;
import com.aspose.words.DmlRenderingMode;
import com.aspose.words.Document;
import com.aspose.words.DocumentProperty;
import com.aspose.words.DrawingML;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.IImageData;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.ImageSize;
import com.aspose.words.ImageType;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.SaveFormat;
import com.aspose.words.Shape;
import com.aspose.words.ShapeRenderer;

public class Emf2 {

	public enum AsposeDrawingType {
		DRAWING_ML(NodeType.DRAWING_ML),
		SHAPE(NodeType.SHAPE);
		private final int code;
		AsposeDrawingType (int code) {
			this.code= code;
		}
		public int code() { return code; }
	}

	public enum AposeWordImageType {
		NO_IMAGE(ImageType.NO_IMAGE),   
		UNKNOWN(ImageType.UNKNOWN),
		EMF(ImageType.EMF),
		WMF(ImageType.WMF),
		PICT(ImageType.PICT),   
		JPEG(ImageType.JPEG),   
		PNG(ImageType.PNG),
		BMP(ImageType.BMP),
		ILLEGAL(-1);

		private int typeCode= -1;
		private AposeWordImageType(int typeCode) {
			this.typeCode= typeCode;
		}

		private int code() {
			return typeCode;
		}
		public static AposeWordImageType fromCode(int typeCode) {
			for (AposeWordImageType asposeType : AposeWordImageType.values()) {
				if (asposeType.code() == typeCode) {
					return asposeType;
				}
			}
			return ILLEGAL;
		}
	}

	public enum AwtImageType {
		INT_RGB(1),
		INT_ARGB(2),
		INT_ARGB_PRE(3),
		INT_BGR(4),
		THREE_BYTE_BGR(5),
		FOUR_BYTE_ABGR(6),
		FOUR_BYTE_ABGR_PRE(7),
		USHORT_565_RGB(8),
		USHORT_555_RGB(9),
		BYTE_GRAY(10),
		USHORT_GRAY(11),
		BYTE_BINARY(12),
		BYTE_INDEXED(13),
		CUSTOM(0),
		ILLEGAL(-1);

		private AwtImageType(int typeCode) {
			this.typeCode= typeCode;
		}

		public int code() {
			return typeCode;
		}
		private int typeCode=0;
		public static AwtImageType fromCode(int typeCode) {
			for (AwtImageType awt : AwtImageType.values()) {
				if (awt.code() == typeCode) {
					return awt;
				}
			}
			return ILLEGAL;
		}

	}
	public static final Logger logger= LoggerFactory.getLogger("extraction:image");

	public static void main(String[] args) throws Exception {

		new com.aspose.words.License().setLicense(Emf2.class.getClassLoader().getResourceAsStream("Aspose.Total.Java.lic.xml"));
		logger.trace("Hello World!");

		File fontLocation= new File(".");
		File[] list = fontLocation.listFiles();
		logger.debug("#of files found={}",list!=null?list.length:0);
		for (File file : list) {
			logger.trace("file=",file.getAbsolutePath());
		}

		File[] packagedFonts= fontLocation.listFiles(new FilenameFilter() {
			public boolean accept(File dir, String filename) {
				logger.trace("Now processing...",filename);
				return filename.toLowerCase().endsWith(".ttf");
			}
		});

		logger.debug("#of fonts found={}",packagedFonts!=null?packagedFonts.length:0);
		for (File fontFile : packagedFonts) {
			Font font= Font.createFont(Font.TRUETYPE_FONT, fontFile);
			boolean success= GraphicsEnvironment.getLocalGraphicsEnvironment().registerFont(font);
			logger.trace("registering font={};success={}",font.getFontName(),success);
		}
		for (String font : java.awt.GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames()) {
			logger.trace("Found font={}",font);
		}
		extractPicturesFromDoc("aspose");
		convertEmfToPng("aspose");

	}
	
	public static void convertEmfToPng(String fileName) throws MetafilesException, FileNotFoundException {
		ImageSaveOptions pngSave= new ImageSaveOptions(com.aspose.words.SaveFormat.PNG);
		pngSave.setResolution(600);
		pngSave.setUseHighQualityRendering(true);
		pngSave.setDmlRenderingMode(DmlRenderingMode.DRAWING_ML);
		pngSave.setDmlEffectsRenderingMode(DmlEffectsRenderingMode.FINE);
		pngSave.setUseAntiAliasing(true);
		EmfMetafileImage emf= new EmfMetafileImage(fileName+".emf");
		emf.save(fileName, pngSave);
	}

	@SuppressWarnings("deprecation")
	public static void extractPicturesFromDoc(String docName) throws Exception {
		Document doc= new Document(docName+".docx");
		Integer emfOrWmfIndex= 1;
		Integer pngOrJpegIndex= 100;
		Integer bmpOrPictIndex= 10000;
		Integer otherIndex= 1000;


		
		String outDir= "out" + File.separator +docName+File.separator;
		FileUtils.forceMkdir(new File(outDir));
		FileWriter html= new FileWriter(outDir+"out.html");          
		html.write("<html>\n<head><meta http-equiv=\"x-ua-compatible\" content=\"IE=edge,chrome=1\"></head><body>\n");

		for (AsposeDrawingType type : AsposeDrawingType.values()) {
			Node[] drawArray= doc.getChildNodes(type.code(), true).toArray();
			int index=0;
			logger.info("type={};count={}",type,drawArray.length);
			for (Node n : drawArray) {
				WordDrawing node=null;
				DrawingML dml= null;
				Shape s= null;
				if (n instanceof Shape) {
					s= (Shape)n;
					node= new WordDrawing(s);
				} else if (n instanceof DrawingML) {
					dml= (DrawingML)n;
					node = new WordDrawing(dml);
				}
										
				index++;
				IImageData img= node.getImageData();
				BufferedImage bi= img.toImage();
				AposeWordImageType asposeWordImageType= AposeWordImageType.fromCode(img.getImageType());
				String extn= null;
				String trimmedDrawingName= node.getName().replace(" ", "")+index;
				ImageSize is= img.getImageSize();
				long resolution= 600;
				int scale= 1000;
				Graphics2D gd= bi.createGraphics();
				gd.getClipBounds();
				int jpegQual= 70;
				boolean antiAlias= true;
				boolean highQualityRendering= true;
				try {
					extn= FileFormatUtil.imageTypeToExtension(img.getImageType());
				} catch (IllegalArgumentException e) {
					extn= "unknown";
				}
				logger.debug("imageType={};name={};imageSize.Width={};imageSize.Height={};"
						+ "imageSize.HorRes={};imageSize.VertRes={};imageSize.WPoints={};imageSize.HPoints={};"
						+ "bufferedImageType={}; biHeight={}; biWidth={}; trimmedDrawingName={}; extn={};"
						+ ""
						+ "bufferedImageInfo={};drawInfo={}", 
						asposeWordImageType,node.getName(), is.getWidthPixels(),is.getHeightPixels(),
						is.getHorizontalResolution(), is.getVerticalResolution(), is.getWidthPoints(),is.getHeightPoints(),
						AwtImageType.fromCode(bi.getType()), bi.getHeight(), bi.getWidth(), trimmedDrawingName, extn,
						bi.toString(), node.toString());
				if (StringUtils.isBlank(node.getName())) {
					if (dml != null) {
						dml.getParentNode();
						logger.debug("getAncestor={}",dml.getAncestor(DocumentProperty.class));
					} else if (s != null) {
						s.getExpandedRunPr_IInline(54);
						
						logger.debug(s.toTxt()+s.getText());
						@SuppressWarnings("unchecked")
						NodeCollection<Node> ns= s.getChildNodes();
						while (ns.iterator().hasNext()) {
							Node n1= (Node)ns.iterator().next();
							n1.getText();
						}
						logger.debug("shape={}",s.getAncestor(DocumentProperty.class));
						s.getParentParagraph();
					}
				}
				if (asposeWordImageType==AposeWordImageType.UNKNOWN) {
					otherIndex++;
					continue;
				}
				if (img==null || asposeWordImageType==AposeWordImageType.NO_IMAGE) {
					continue;
				}
				if (asposeWordImageType==AposeWordImageType.EMF || asposeWordImageType==AposeWordImageType.WMF) {
					
					ShapeRenderer sr= node.getShapeRenderer();
					img.save(outDir+trimmedDrawingName+extn);
					EmfMetafileImage emf= new EmfMetafileImage(outDir+trimmedDrawingName+extn);
					trimmedDrawingName +=  "_" + scale + "_" + resolution + "_" + jpegQual + "_" + antiAlias + "_" + highQualityRendering;
					ImageSaveOptions pngSave= new ImageSaveOptions(com.aspose.words.SaveFormat.PNG);
					pngSave.setResolution(resolution);
					pngSave.setUseHighQualityRendering(highQualityRendering);
					pngSave.setDmlRenderingMode(DmlRenderingMode.DRAWING_ML);
					pngSave.setDmlEffectsRenderingMode(DmlEffectsRenderingMode.FINE);
					pngSave.setUseAntiAliasing(antiAlias);
					pngSave.setScale((float)scale/1000);
					
					PngOptions pngOptions= new PngOptions();
							

					ImageSaveOptions jpgSave= new ImageSaveOptions(SaveFormat.JPEG);
					jpgSave.setUseHighQualityRendering(true);
					jpgSave.setResolution(resolution);
					jpgSave.setJpegQuality(jpegQual);
					jpgSave.setScale((float)scale/1000);


					sr.save(outDir+trimmedDrawingName+".png", pngSave);
					BufferedImage savedPNG= ImageIO.read(new File(outDir+trimmedDrawingName+".png"));
					BufferedImage resizedFromSaved= Scalr.resize(savedPNG, Method.ULTRA_QUALITY, Mode.FIT_TO_WIDTH, 435);
					BufferedImage resizedFromBi= Scalr.resize(bi, Method.ULTRA_QUALITY, Mode.FIT_TO_WIDTH, 435);
					emf.save(outDir+trimmedDrawingName+"_buffered_emf.png", pngOptions);
					ImageIO.write(bi, "png", new File(outDir+trimmedDrawingName+"_buffered.png"));
					ImageIO.write(resizedFromSaved, "png", new File(outDir+trimmedDrawingName+"_resized_from_saved_scalr_antialias_435.png"));
					ImageIO.write(resizedFromBi, "png", new File(outDir+trimmedDrawingName+"_resized_from_bi_scalr_antialias_435.png"));
					//sr.save(outDir+trimmedDrawingName+".jpg", jpgSave);

					html.write("\t<div>\n\t\t\n\t\t<br>\n\t\t<hr><p align=center>.SVG figure: "+ trimmedDrawingName + 
							"</p>\n\t\t<hr>\n\t\t<br>\n\t\t<br>\n\t\t<img src=\""+trimmedDrawingName+".svg\" width=\"100%\" />\n\t</div>\n");

					//convertToSVG(outputDir + docId + "\\", trimmedDrawingName, extn);
					emfOrWmfIndex++;
				} else if (asposeWordImageType==AposeWordImageType.PNG || asposeWordImageType==AposeWordImageType.JPEG) {
					ShapeRenderer sr= node.getShapeRenderer();
					ImageSaveOptions pngSave= new ImageSaveOptions(com.aspose.words.SaveFormat.PNG);
					pngSave.setResolution(resolution);
					pngSave.setUseHighQualityRendering(highQualityRendering);
					pngSave.setDmlRenderingMode(DmlRenderingMode.DRAWING_ML);
					pngSave.setDmlEffectsRenderingMode(DmlEffectsRenderingMode.FINE);
					pngSave.setUseAntiAliasing(antiAlias);
					pngSave.setScale((float)scale/1000);
					img.save(outDir+trimmedDrawingName+extn);
					sr.save(outDir+trimmedDrawingName+"_DIRECT"+extn, pngSave);
					if (is.getHeightPoints()>99) {
						html.write("\t<div>\n\t\t\n\t\t<br>\n\t\t<hr><p align=center>"+extn.toUpperCase()+" figure: "+ trimmedDrawingName + "</p>\n\t\t<hr>\n\t\t<br>\n\t\t<br>\n\t\t<img src=\""+trimmedDrawingName+extn+"\" width=\"100%\" />\n\t</div>\n");
					}
					pngOrJpegIndex++;
				} else if (asposeWordImageType==AposeWordImageType.BMP || asposeWordImageType==AposeWordImageType.PICT){
					img.save(outDir+bmpOrPictIndex + extn);
					bmpOrPictIndex++;
				} else {
					logger.info("PICT type={}; isLink={}; isLinkOnly={}; imageSize={}; sourceFileName={}; hasImage={}",
							asposeWordImageType, img.isLink(), img.isLinkOnly(), img.getImageSize().getHorizontalResolution(), img.getSourceFullName(), img.hasImage());
				}
			}
		}
		html.write("</body>\n</html>");
		html.close();
	}
}
