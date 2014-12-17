package com.xauto.ux;

import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.GraphicsEnvironment;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileWriter;
import java.io.FilenameFilter;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.imgscalr.Scalr;
import org.imgscalr.Scalr.Method;
import org.imgscalr.Scalr.Mode;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.aspose.imaging.fileformats.metafile.EmfMetafileImage;
import com.aspose.imaging.fileformats.metafile.WmfMetafileImage;
import com.aspose.imaging.imageoptions.PngOptions;
import com.aspose.words.DmlEffectsRenderingMode;
import com.aspose.words.DmlRenderingMode;
import com.aspose.words.Document;
import com.aspose.words.DrawingML;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.IImageData;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.ImageSize;
import com.aspose.words.ImageType;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.SaveFormat;
import com.aspose.words.Shape;
import com.aspose.words.ShapeRenderer;

public class Emf2 {

	public enum AsposeDrawingType {
		DRAWING_ML(NodeType.DRAWING_ML),
		SHAPE(NodeType.SHAPE),
		GROUP_SHAPE(NodeType.GROUP_SHAPE);
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
			logger.debug("registering font={};success={}",font.getFontName(),success);
		}
		for (String font : java.awt.GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames()) {
			logger.debug("Found font={}",font);
		}
		extractPicturesFromDoc("10008965");
		BinaryPart bp;

	}

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
			List<WordDrawing> drawings= new ArrayList<WordDrawing>();
			for (Node node : drawArray) {
				if (node instanceof DrawingML) {
					drawings.add(new WordDrawing((DrawingML)node));
				} else if (node instanceof Shape) {
					drawings.add(new WordDrawing((Shape)node));
				}
			}
			int index=0;
			logger.info("type={};count={}",type,drawings.size());
			for (WordDrawing node : drawings) {
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
				
				String[] properties= bi.getPropertyNames();
				logger.debug("imageType={};name={};hasImage()={};imageByteSize={};isLink={};imageSize.Width={};imageSize.Height={};"
						+ "imageSize.HorRes={};imageSize.VertRes={};imageSize.WPoints={};imageSize.HPoints={};"
						+ "bufferedImageType={}; biHeight={}; biWidth={}; trimmedDrawingName={}; extn={};"
						+ ""
						+ "bufferedImageInfo={};drawInfo={}", 
						asposeWordImageType,node.getName(), img.hasImage(), img.getImageBytes()==null?0:img.getImageBytes().length, img.isLink(),is.getWidthPixels(),is.getHeightPixels(),
						is.getHorizontalResolution(), is.getVerticalResolution(), is.getWidthPoints(),is.getHeightPoints(),
						AwtImageType.fromCode(bi.getType()), bi.getHeight(), bi.getWidth(), trimmedDrawingName, extn,
						bi.toString(), node.toString());
				StringBuilder sb= new StringBuilder();
				if (properties != null) {
					for (String s : properties) {
					sb.append(s).append(";");
					}
				}
				logger.debug("\t\tproperties={}",sb.toString());
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
					PngOptions pngOptions= new PngOptions();
					if (asposeWordImageType==AposeWordImageType.EMF) {
						EmfMetafileImage emf= new EmfMetafileImage(outDir+trimmedDrawingName+extn);
						emf.save(outDir+trimmedDrawingName+"_buffered_emf.png", pngOptions);
					} else {
						WmfMetafileImage wmf= new WmfMetafileImage(outDir+trimmedDrawingName+extn);
						wmf.save(outDir+trimmedDrawingName+"_buffered_emf.png", pngOptions);
					}
						
					trimmedDrawingName +=  "_" + scale + "_" + resolution + "_" + jpegQual + "_" + antiAlias + "_" + highQualityRendering;
					ImageSaveOptions pngSave= new ImageSaveOptions(com.aspose.words.SaveFormat.PNG);
					pngSave.setResolution(resolution);
					pngSave.setUseHighQualityRendering(highQualityRendering);
					pngSave.setDmlRenderingMode(DmlRenderingMode.DRAWING_ML);
					pngSave.setDmlEffectsRenderingMode(DmlEffectsRenderingMode.FINE);
					pngSave.setUseAntiAliasing(antiAlias);
					pngSave.setScale((float)scale/1000);
					
							

					ImageSaveOptions jpgSave= new ImageSaveOptions(SaveFormat.JPEG);
					jpgSave.setUseHighQualityRendering(true);
					jpgSave.setResolution(resolution);
					jpgSave.setJpegQuality(jpegQual);
					jpgSave.setScale((float)scale/1000);


					sr.save(outDir+trimmedDrawingName+".png", pngSave);
					BufferedImage savedPNG= ImageIO.read(new File(outDir+trimmedDrawingName+".png"));
					BufferedImage resizedFromSaved= Scalr.resize(savedPNG, Method.ULTRA_QUALITY, Mode.FIT_TO_WIDTH, 435);
					BufferedImage resizedFromBi= Scalr.resize(bi, Method.ULTRA_QUALITY, Mode.FIT_TO_WIDTH, 435);
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
