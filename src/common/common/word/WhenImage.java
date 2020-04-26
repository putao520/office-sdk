package common.common.word;

@FunctionalInterface
public interface WhenImage {
    /**
     * @apiNote 当遇到图片处理时
     * @param keyName 域名称
     * @param value 当前值
     * */
    byte[] when(String keyName,String value);
}
