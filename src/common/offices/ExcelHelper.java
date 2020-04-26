/**
 * 
 */
/**
 * @author Administrator
 *
 */
package common.offices;

import com.aspose.cells.License;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import common.java.nlogger.nlogger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

public class ExcelHelper {
	private File file;

	public ExcelHelper(String xlsPath) {
		file = new File(xlsPath);
	}

	public static boolean getLicense() throws Exception {
		boolean result = false;
		try {
			InputStream is = com.aspose.cells.License.class.getResourceAsStream("/key.xml");
			License aposeLic = new License();
			aposeLic.setLicense(is);
			result = true;
			is.close();
		} catch (Exception e) {
			nlogger.logInfo(e);
			throw e;
		}
		return result;
	}

	public boolean toHtml(String htmlPath) {
		boolean r = true;
		try {
			if (!getLicense()) { // 验证License 若不验证则转化出的pdf文档有水印
				throw new Exception("com.aspose.cell lic ERROR!");
			}
			System.out.println(file.getAbsolutePath() + " -> " + htmlPath);
			new Workbook(file.getAbsolutePath()).save(htmlPath, SaveFormat.HTML);
		} catch (Exception e) {
			nlogger.logInfo(e);
			r = false;
		}
		return r;
	}


	/**根据请求String的返回值，生成excel临时文件，然后提供下载
	 * @param reqString
	 * @return
	 * @throws IOException 
	 */
	@SuppressWarnings("unchecked")
	public final static byte[] out(String reqString) throws IOException{
		ByteArrayOutputStream outByte = new ByteArrayOutputStream();
		JSONArray array = null;
		String rlt = reqString;
		/*
		String rlt;
		array = new JSONArray();
		JSONObject object = new JSONObject();
		object.put("1测试1", "数值1");
		object.put("1测试2", "数值2");
		array.add(object);
		object = new JSONObject();
		object.put("2测试1", "数值1");
		object.put("2测试2", "数值2");
		array.add(object);
		rlt = array.toJSONString();
		*/
        Object _obj = JSONArray.toJSONArray(rlt);
		if( _obj == null ){
            _obj = JSONObject.toJSON(rlt);
			if( _obj != null){
				array = new JSONArray();
                array.add(_obj);
			}
		}
		else{
			array = (JSONArray)_obj;
		}
		
		if( array != null){
			HSSFWorkbook excel = jsonArray2Excel(array);
			excel.write(outByte);
		}
		return outByte.toByteArray();
	}
	private static HSSFWorkbook jsonArray2Excel(JSONArray ary){
		HSSFWorkbook excel = new HSSFWorkbook();
		HSSFSheet sheet = excel.createSheet();
		HSSFRow titleLine = sheet.createRow(0);
		JSONObject jsonObject = (JSONObject) ary.get(0);
		HSSFCell cell;
		int i = 0;
		for( Object keyname : jsonObject.keySet()){
			cell = titleLine.createCell(i);
			cell.setCellValue(keyname.toString());
			i++;
		}
		int l = 1;
		HSSFRow dataLine;
		for(Object _obj : ary){
			i = 0;
			dataLine = sheet.createRow(l);
			jsonObject = (JSONObject)_obj;
			for( Object _item :jsonObject.keySet() ){
				cell = dataLine.createCell(i);
				cell = setvalue(cell, jsonObject.get(_item));
				i++;
			}
			l++;
		}
		return excel;
	}
	private static HSSFCell setvalue(HSSFCell cell,Object value){
		if( value == null ){
			cell.setCellValue( "" );
		}
		else{
			if( value instanceof String ){
				cell.setCellValue((String)value);
			}
			else if( value instanceof Boolean ){
				cell.setCellValue(Boolean.parseBoolean(value.toString()));
			}
			else if( value instanceof Double ){
				cell.setCellValue(Double.parseDouble(value.toString()));
			}
			else if( value instanceof Date ){
				cell.setCellValue( (Date)value );
			}
			else {
				cell.setCellValue( value.toString() );
			}
		}
		return cell;
	}
}