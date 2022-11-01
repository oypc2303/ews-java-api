package microsoft.exchange.webservices.data.exchangeservice;

import lombok.Data;

/**
 * @author oypc2
 * @version 1.0.0
 * @title MailContentVO
 * @date 2022年10月17日 15:47:26
 * @description TODO
 */
@Data
public class MailContentVO {

    /**
     * 正文信息类型
     */
    private MailContentType type;

    /**
     * 正文信息
     * txt内容或img绝对路径
     */
    private String content;

}
