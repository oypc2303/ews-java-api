package microsoft.exchange.webservices.testApi;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.List;

/**
 * https://blog.csdn.net/u013943849/article/details/115543469
 * @author oypc2
 * @version 1.0.0
 * @title TestApi
 * @date 2022年10月14日 15:06:44
 * @description TODO
 */
public class TestApi {
    String domain = "mytest.com";
    public final static String username= "ouyangpengcheng";
    public final static String password= "111111";

    public static void main(String[] args) throws Exception {
        new TestApi().send("nihao", "123123");
    }

    private ExchangeService getExchangeService() {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
        //用户认证信息
        ExchangeCredentials credentials;
        if (domain == null) {
            credentials = new WebCredentials(username, password);
        } else {
            credentials = new WebCredentials(username, password, domain);
        }
        service.setCredentials(credentials);
        try {
//            service.setUrl(new URI("https://exchange.mytest.com/ecp"));
            service.setUrl(new URI("https://exchange.mytest.com/EWS/Exchange.asmx"));
        } catch (URISyntaxException e) {
            e.printStackTrace();
        }
        return service;
    }

    public void send(String subject, String bodyText) throws Exception {
        ExchangeService service = getExchangeService();
        EmailMessage msg = new EmailMessage(service);
        msg.setSubject(subject);
        MessageBody body = MessageBody.getMessageBodyFromText(bodyText);
//        body.setBodyType(BodyType.HTML);
        msg.setBody(body);
        msg.getToRecipients().add("ouyangpengcheng@mytest.com");
        msg.getAttachments().addFileAttachment("D:/test.txt");
//        msg.send();
        msg.sendAndSaveCopy();
    }
}
