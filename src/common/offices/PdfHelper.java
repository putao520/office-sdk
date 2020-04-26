package common.offices;


import com.aspose.pdf.*;
import com.aspose.pdf.operators.ConcatenateMatrix;
import com.aspose.pdf.operators.Do;
import com.aspose.pdf.operators.GRestore;
import com.aspose.pdf.operators.GSave;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;

public class PdfHelper {
    private Document doc;
    public static final PdfHelper open(String docPath){
        return new PdfHelper(docPath);
    }
    public static final PdfHelper open(File docFile){
        return new PdfHelper(docFile);
    }
    public static final PdfHelper open(InputStream docStream){
        return new PdfHelper((docStream));
    }
    private PdfHelper(String pdfFile) {
        PdfHelper.getLicense();
        doc = new Document(pdfFile);
    }
    private PdfHelper(File pdfFile) {
        PdfHelper.getLicense();
        doc = new Document(pdfFile.getAbsolutePath());
    }
    private PdfHelper(InputStream pdfFile) {
        PdfHelper.getLicense();
        doc = new Document(pdfFile);
    }
    public static boolean getLicense(){
        boolean result = false;
        try {
            InputStream is = com.aspose.pdf.Document.class.getResourceAsStream("/key.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("com.aspose.pdf lic ERROR!");
        }
        return result;
    }
    public PdfHelper convert(int format){
        doc.convert("file.log", format, ConvertErrorAction.Delete);
        return this;
    }

    public PdfHelper toHtml(String htmlPath) {
        try {
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            doc.save(htmlPath, saveOptions);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return this;
    }
    public PdfHelper toDoc(String docPath) {
        try {
            doc.save(docPath, SaveFormat.DocX);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return this;
    }
    public PdfHelper insertBookmark(String text){
        OutlineCollection outLine = doc.getOutlines();
        OutlineItemCollection pdfOutline = new OutlineItemCollection(outLine);
        pdfOutline.setTitle(text);
        pdfOutline.setItalic(true);
        pdfOutline.setBold(true);
        //set the destination page number
        pdfOutline.setAction( new GoToAction(doc.getPages().get_Item(1)));
        //add bookmark in the document's outline collection.
        outLine.add(pdfOutline);
        doc.save();
        //save output
        return this;
    }
    public PdfHelper addWatermark(String text){
        // create text stamp
        TextStamp textStamp = new TextStamp(text);
        // set whether stamp is background
        textStamp.setBackground(true);
        // set origin
        textStamp.setXIndent(100);
        textStamp.setYIndent(100);
        // rotate stamp
        textStamp.setRotate(Rotation.on90);
        // set text properties
        textStamp.getTextState().setFont(new FontRepository().findFont("Arial"));
        textStamp.getTextState().setFontSize(14.0F);
        textStamp.getTextState().setFontStyle(FontStyles.Bold);
        textStamp.getTextState().setFontStyle(FontStyles.Italic);
        textStamp.getTextState().setForegroundColor(Color.getGray());

        PageCollection pages = doc.getPages();
		// ExStart:InfoClass
		// iterate through all pages of PDF file
		for (int Page_counter = 1,Page_Max = doc.getPages().size(); Page_counter <= Page_Max; Page_counter++) {
			// add stamp to all pages of PDF file
            pages.get_Item(Page_counter).addStamp(textStamp);
		}
		// ExEnd:InfoClass
        doc.save();
        return this;
    }
    public PdfHelper appendBefore(InputStream imageSteam, Rectangle rectangle){
        Page page = doc.getPages().get_Item(1);
        // 创建Image对象，命名空间是必要的，因为在别的命名空间也有Image类
        XImageCollection xImages = page.getResources().getImages();
        xImages.add(imageSteam);
        XImage ximage = xImages.get_Item(page.getResources().getImages().size());
        // Rectangle rectangle = new Rectangle(lowerLeftX, lowerLeftY, upperRightX, upperRightY);
        Matrix matrix = new Matrix(new double[] { rectangle.getURX() - rectangle.getLLX(), 0, 0, rectangle.getURY() - rectangle.getLLY(), rectangle.getLLX(), rectangle.getLLY() });
        Operator[] opArray = {
                new GSave(),
                new ConcatenateMatrix(matrix),
                new Do(ximage.getName()),
                new GRestore()
        };
        page.getContents().add(opArray);
        return this;
    }
    public byte[] toBytes(){
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        doc.save(out);
        byte[] o = out.toByteArray();
        try{
            out.close();
        }
        catch (Exception e){}
        return o;
    }
    public PdfHelper save(){
        doc.save();
        return this;
    }
}
