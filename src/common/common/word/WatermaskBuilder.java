package common.common.word;

import com.aspose.words.Shape;

@FunctionalInterface
public interface WatermaskBuilder {
    Shape[] run(WatermarkText watermarkText);
}
