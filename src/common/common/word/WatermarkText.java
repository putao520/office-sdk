package common.common.word;


import java.awt.*;

public class WatermarkText {
    private int width = 0;
    private int height= 25;
    private String text;
    private int rotation = 0;
    private java.awt.Color color = WatermarkText.str2Color("E0E0E0");
    private static final java.awt.Color str2Color(String colorStr){
        return new java.awt.Color(Integer.parseInt(colorStr, 16));
    }

    public static final WatermarkText build(String text){
        return new WatermarkText(text);
    }
    private WatermarkText(String text){
        this.text = text;
    }

    public int getWidth() {
        return this.width > 0 ? width : (this.height == 0 ? 25 : this.height) * this.text.length();
    }

    public WatermarkText setWidth(int width) {
        this.width = width;
        return this;
    }

    public int getHeight() {
        return height;
    }

    public WatermarkText setHeight(int height) {
        this.height = height;
        return this;
    }

    public String getText() {
        return text;
    }

    public WatermarkText setText(String text) {
        this.text = text;
        return this;
    }

    public int getRotation() {
        return rotation;
    }

    public WatermarkText setRotation(int rotation) {
        this.rotation = rotation;
        return this;
    }

    public Color getColor() {
        return color;
    }

    public WatermarkText setColor(String color) {
        this.color = WatermarkText.str2Color( color );
        return this;
    }
}
