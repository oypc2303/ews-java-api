package microsoft.exchange.webservices.data.exchangeservice;

/**
 * @author oypc2
 * @version 1.0.0
 * @title MailContentType
 * @date 2022年10月17日 15:47:11
 * @description TODO
 */
public enum MailContentType {
    txt("文本"),img("图片");
    private String type;
    MailContentType(String type) {
        this.type = type;
    }
}
