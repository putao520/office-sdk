package common.offices;

import com.aspose.words.*;
import common.common.word.WatermarkText;
import common.common.word.WatermaskBuilder;
import common.common.word.WatermaskFilter;
import common.common.word.WhenImage;
import common.java.nlogger.nlogger;
import org.json.simple.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;

public class WordHelper {
	public enum  Align{
		TOP,CENTER,BOTTOM,FULL
	}
	private static String normalCharSet;
	private Document doc;
	static{
		normalCharSet = "utf8";
		getLicense();
	}
	public static final WordHelper open(String docPath){
		return new WordHelper(docPath);
	}
	public static final WordHelper open(File docFile){
		return new WordHelper(docFile);
	}
	public static final WordHelper open(InputStream docStream){
		return new WordHelper((docStream));
	}
	private WordHelper(String docPath) {
		try{
			this.doc = new Document(docPath);
		}
		catch (Exception e){
			nlogger.logInfo(e);
		}
	}
	private WordHelper(File docFile) {
		try{
			this.doc = new Document(docFile.getAbsolutePath());
		}
		catch (Exception e){
			nlogger.logInfo(e);
		}
	}
	private WordHelper(InputStream docStream) {
		try{
			this.doc = new Document(docStream);
		}
		catch (Exception e){
			nlogger.logInfo(e);
		}
	}

	public static boolean getLicense(){
		boolean result = false;
		try {
			InputStream is = com.aspose.words.Document.class.getResourceAsStream("/key.xml");
			License aposeLic = new License();
			aposeLic.setLicense(is);
			result = true;
			is.close();
		} catch (Exception e) {
			nlogger.logInfo(e);
		}
		return result;
	}

	/**
	 * @apiNote 插入文字水印
	 * @param text 文字内容
	 * @param ag   定位设置
	 * */
	public WordHelper insertWatermarkText(WatermarkText text,Align ag) {
		int[] hfType = new int[]{HeaderFooterType.HEADER_PRIMARY};
		WatermaskBuilder wmb = null;
		WatermaskFilter wmf = null;
		return insertWatermarkText(text, ag, hfType, wmb, wmf);
	}
	public WordHelper insertWatermarkText(WatermarkText text,Align ag, int[] HeaderFooterType, WatermaskBuilder builder, WatermaskFilter filter){
		switch (ag){
			case TOP:
				insertWatermarkText(text,HeaderFooterType, builder, (i,watermark)->{
					watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.MARGIN);
					watermark.setRelativeVerticalPosition(RelativeVerticalPosition.MARGIN);
					watermark.setWrapType(WrapType.NONE);
					//  我们需要自定义距离顶部的高度
					// watermark.setVerticalAlignment(VerticalAlignment.TOP);
					watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
					// 设置距离顶部的高度
					watermark.setTop(160);
					return watermark;
				});
				break;
			case CENTER:
				insertWatermarkText(text,HeaderFooterType, builder, (i,watermark)->{
					watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
					watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
					watermark.setWrapType(WrapType.NONE); // TOP_BOTTOM : 将所设置位置的内容往上下顶出去
					watermark.setVerticalAlignment(VerticalAlignment.CENTER);
					watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
					return watermark;
				});
				break;
			case BOTTOM:
				insertWatermarkText(text,HeaderFooterType, builder, (i,watermark)->{
					watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.MARGIN);
					watermark.setRelativeVerticalPosition(RelativeVerticalPosition.MARGIN);
					watermark.setWrapType(WrapType.NONE);
					// 我们需要自定义距离顶部的高度
					// watermark.setVerticalAlignment(VerticalAlignment.BOTTOM);
					watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
					// 设置距离顶部的高度
					watermark.setTop(480);
					return watermark;
				});
				break;
			case FULL:
				insertWatermarkIntoPage(text);
		}
		return this;
	}

	private int _HorizontalAlignment = RelativeHorizontalPosition.DEFAULT;
	private double fixX= 0.0;
	private int _VerticalAlignment = RelativeVerticalPosition.TEXT_FRAME_DEFAULT;
	private double fixY= 0.0;
	public WordHelper WatermarkImagePostion(int HorizontalAlignment, double fixX, int VerticalAlignment, double fixY){
		this._HorizontalAlignment = HorizontalAlignment;
		this._VerticalAlignment = VerticalAlignment;
		this.fixX = fixX;
		this.fixY = fixY;
		return this;
	}
	/**
	 * @apiNote 插入图片水印到指定书签位置
	 * @param imageFileStream 图片数据流
	 * @param bookmarkName  书签名称
	 * */
	public WordHelper insertWatermarkImage(InputStream imageFileStream,String bookmarkName){
		Shape watermark = new Shape(doc, ShapeType.IMAGE);
		try{
			watermark.getImageData().setImage( imageFileStream );
			insertWatermarkImage(watermark, bookmarkName);
		}
		catch (Exception e){
			nlogger.logInfo("水印文件 ->不存在");
		}
		return this;
	}
	public WordHelper insertWatermarkImage(String imageFilePath,String bookmarkName){
		Shape watermark = new Shape(doc, ShapeType.IMAGE);
		try{
			watermark.getImageData().setImage( imageFilePath );
			insertWatermarkImage(watermark, bookmarkName);
		}
		catch (Exception e){
			nlogger.logInfo("水印文件[" + imageFilePath + "] ->不存在");
		}
		return this;
	}
	private void insertWatermarkImage(Shape watermark,String bookmarkName){
		watermark.setWrapType( WrapType.NONE );
		watermark.setBehindText(true);
		try {
			// 水印大小
			double height = 70.0;
			watermark.setWidth(70);
			watermark.setHeight(height);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		DocumentBuilder docBuilder = new DocumentBuilder(doc);
		try{
			if( !docBuilder.moveToBookmark(bookmarkName) ){
				nlogger.logInfo("书签[" + bookmarkName + "]不存在");
			}
			// 设置水印定位
			watermark.setRelativeHorizontalPosition( RelativeHorizontalPosition.COLUMN );
			watermark.setHorizontalAlignment( _HorizontalAlignment );
			watermark.setLeft( fixX );
			watermark.setRelativeVerticalPosition( RelativeVerticalPosition.LINE );
			watermark.setVerticalAlignment( _VerticalAlignment );
			watermark.setTop( fixY );

			// 插入水印
			docBuilder.insertNode(watermark);
			// 删除定位用书签
			doc.getRange().getBookmarks().get(bookmarkName).remove();
		}
		catch (Exception e){
			nlogger.logInfo("书签[" + bookmarkName + "]不存在");
		}
	}

	private void insertWatermarkIntoPage(WatermarkText watermarkText){
		final int lineNum = 4;
		final int rowNum = 3;
		int[] hfType = new int[]{HeaderFooterType.HEADER_FIRST,HeaderFooterType.HEADER_PRIMARY,HeaderFooterType.HEADER_EVEN};
		WatermaskBuilder watermaskBuilder = wmt->{
			List<Shape> shapeArray = new ArrayList<>();
			try{
				Document tDoc = doc.deepClone();
				for( int i=0,l = tDoc.getPageCount(); i<l;i++ ){
					PageInfo pInfo = tDoc.getPageInfo(i);
					int pH = (int)pInfo.getHeightInPoints();
					int pW = (int)pInfo.getWidthInPoints();
					int startTop = (pH - (watermarkText.getHeight() * lineNum))/lineNum;
					int startLeft= (pW - (watermarkText.getWidth()  * rowNum ))/rowNum;
					for(int c =0; c <lineNum; c++){
						for(int b =0; b <rowNum; b++){
							Shape shape = buildTextShape(wmt);
							// 设置页相对顶边距
							shape.setTop( (startTop + watermarkText.getHeight()) * c );
							// 设置页相对左边距
							shape.setLeft((startLeft + watermarkText.getWidth()) * b );
							// 添加到图形数组
							shapeArray.add(shape);
						}
					}
				}
				tDoc.cleanup();
			}
			catch (Exception e){
				e.printStackTrace();
			}
			return shapeArray.toArray(new Shape[shapeArray.size()]);
		};
		insertWatermarkText(watermarkText, hfType, watermaskBuilder, null);
	}

	public int getPageSize(){
		try{
			return doc.getPageCount();
		}
		catch (Exception e){
			return 1;
		}
	}

	private Shape buildTextShape(WatermarkText watermarkText){
		Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
		watermark.getTextPath().setText(watermarkText.getText());
		watermark.getTextPath().setFontFamily("宋体");//Arial;
		try {
			// 水印大小
			watermark.setWidth(watermarkText.getWidth());
			watermark.setHeight(watermarkText.getHeight());
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		// 左下到右上
		watermark.setRotation(watermarkText.getRotation());
		watermark.getFill().setColor(watermarkText.getColor()); // Try Color.lightGray to get more Word-style watermark
		watermark.setStrokeColor(watermarkText.getColor()); // Try Color.lightGray to get more Word-style watermark
		watermark.setWrapType(WrapType.NONE);
		return watermark;
	}

	private void insertWatermarkText(Shape[] watermark, int[] headerTypeArray, WatermaskFilter watermaskPositionConfigFunc)  {
		Paragraph watermarkPara = new Paragraph(doc);
		for(int i =0,l =watermark.length; i <l; i++){
			watermarkPara.appendChild(watermaskPositionConfigFunc.run(i, watermark[i]));
		}
		for (Section sect : doc.getSections()) {
			for( int i =0,l =headerTypeArray.length; i<l; i++ ){
				insertWatermarkIntoHeader(watermarkPara, sect, headerTypeArray[i]);
			}
		}
	}

	private void insertWatermarkText(WatermarkText watermarkText, int[] headerTypeArray, WatermaskBuilder wmBuilder, WatermaskFilter wmFilter)  {
		if( wmBuilder == null ){
			wmBuilder = wmt->new Shape[]{buildTextShape(wmt)};
		}
		if( wmFilter == null ){
			wmFilter = (i,wmf)->wmf;
		}
		insertWatermarkText(wmBuilder.run( watermarkText ), headerTypeArray, wmFilter);
	}

	private static void insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, int headerType) {
		HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);
		if (header == null) {
			// There is no header of the specified type in the current section, create it.
			header = new HeaderFooter(sect.getDocument(), headerType);
			sect.getHeadersFooters().add(header);
		}


		// Insert a clone of the watermark into the header.
		try {
			header.appendChild(watermarkPara.deepClone(true));
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	public WordHelper toHTML(String htmlPath) {
		try {
			// System.out.println(" -> " + htmlPath);
			doc.save(htmlPath, SaveFormat.HTML);
		} catch (Exception e) {
			nlogger.logInfo(e);
		}
		return this;
	}
	public WordHelper toPdf(String pdfPath) {
		try {
			// System.out.println(" -> " + pdfPath);
			doc.save(pdfPath, SaveFormat.PDF);
		} catch (Exception e) {
			nlogger.logInfo(e);
		}
		return this;
	}

	public WordHelper toPdf(OutputStream pdfStream) {
		try {
			doc.save(pdfStream, SaveFormat.PDF);
		} catch (Exception e) {
			nlogger.logInfo(e);
		}
		return this;
	}
	public byte[] toBytes(int SaveFormat){
		return this.toBytes(SaveFormat, null);
	}
	public byte[] toBytes(int SaveFormat, IImageSavingCallback whenImageWillSave){
		byte[] o = null;
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		try {
			switch (SaveFormat){
				case com.aspose.words.SaveFormat.HTML:
					HtmlSaveOptions hso = new HtmlSaveOptions();
					hso.setEncoding(Charset.forName("UTF-8"));
					if( whenImageWillSave != null ){
						hso.setImageSavingCallback( whenImageWillSave);
					}
					else {
						hso.setExportImagesAsBase64(true);
					}
					doc.save(out, hso);
					break;
				default:
					doc.save(out, SaveFormat);
					break;
			}
			o = out.toByteArray();
		} catch (Exception e) {
			nlogger.logInfo(e);
		}
		finally {
			try {
				out.close();
			}
			catch (Exception e){}
		}
		return o;
	}

	/**
	 * 渲染doc模板文件
	 * */
	public WordHelper render(JSONObject inputData){
		return render(inputData, null);
	}
	public WordHelper render(JSONObject inputData, WhenImage wImage){
		try {
			// Document doc = new Document(file.getAbsolutePath());
			List<String> keyList = new ArrayList<>();
			List<Object> valList = new ArrayList<>();
			for( String key : inputData.keySet() ){
				keyList.add(key);
				valList.add(inputData.get(key));
			}
			MailMerge mm = doc.getMailMerge();

			mm.setFieldMergingCallback(new IFieldMergingCallback() {
				@Override
				public void fieldMerging(FieldMergingArgs fieldMergingArgs) throws Exception {
					/*
					DocumentBuilder builder = new DocumentBuilder(fieldMergingArgs.getDocument());
					builder.moveToMergeField(fieldMergingArgs.getFieldName());
					*/
					// builder.insertField(fieldMergingArgs.getFieldValue() );
				}

				@Override
				public void imageFieldMerging(ImageFieldMergingArgs imageFieldMergingArgs) throws Exception {
					String val = "空图片";
					try{
						val = imageFieldMergingArgs.getFieldValue().toString();
						DocumentBuilder builder = new DocumentBuilder(imageFieldMergingArgs.getDocument());
						builder.moveToMergeField(imageFieldMergingArgs.getFieldName());
						if( wImage != null ){
							builder.insertImage(wImage.when(imageFieldMergingArgs.getFieldName(), val) );
						}
						else {
							builder.insertImage(imageFieldMergingArgs.getFieldValue().toString());
						}
					}
					catch (Exception e){
						System.out.println("渲染图片:[" + val + "] ->失败");
					}
				}
			});


			String[] keys = keyList.toArray(new String[keyList.size()]);
			Object[] vals = valList.toArray(new Object[valList.size()]);
			mm.execute(keys, vals);
			// doc.save(file.getAbsolutePath());
		} catch (Exception e) {
			nlogger.logInfo(e);
		}
		return this;
	}

	private void chkLicense(){
		try{
			if (!getLicense()) { // 验证License 若不验证则转化出的文档有水印
				throw new Exception("com.aspose.words lic ERROR!");
			}
		}
		catch (Exception e){
			nlogger.logInfo(e);
		}
	}

}
