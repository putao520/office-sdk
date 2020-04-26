package test.java;

import common.common.word.WatermarkText;
import common.offices.WordHelper;
import org.json.simple.JSONObject;

public class main {
    public static void main( String[] args ) {
        main.test();
    }
    public static void test(){
        JSONObject data = JSONObject.putx("test","文字模板")
                .puts("pic","e:\\\\test\\1");
        WordHelper
                .open("e:\\\\test\\test.docx")
                .render(data)
                .insertWatermarkText(WatermarkText.build("测试水印").setRotation(-40), WordHelper.Align.FULL)
                // .toBytes(SaveFormat.HTML);
                .toPdf("e:\\\\test\\5.pdf");
    }
}