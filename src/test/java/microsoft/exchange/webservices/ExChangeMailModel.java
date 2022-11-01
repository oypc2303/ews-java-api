package microsoft.exchange.webservices;

import lombok.Data;

import java.util.Date;

/**
 * https://blog.csdn.net/Sayonara_LM/article/details/107711147
 * @author oypc2
 * @version 1.0.0
 * @title ExChangeMailModel
 * @date 2022年10月18日 16:27:03
 * @description TODO
 */
@Data
public class ExChangeMailModel {
    private String id;
    private String subject;
    private Date dateTimeReceived;
    private Boolean isRead;
    private String body;
}
