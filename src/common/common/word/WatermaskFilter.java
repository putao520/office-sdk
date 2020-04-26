package common.common.word;

import com.aspose.words.Shape;

@FunctionalInterface
public interface WatermaskFilter {
    Shape run(int idx, Shape shape);
}
