package microsoft.exchange.webservices.exchange;

import java.io.File;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.List;
import lombok.extern.slf4j.Slf4j;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.WebProxy;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

/**
 * @author oypc2
 * @version 1.0.0
 * @title ExchangeMailUtil
 * @date 2022年10月14日 13:37:18
 * @description TODO
 */
@Slf4j
public class ExchangeMailUtil {
    private String mailServer;
    private String user;
    private String password;
    String to;

    public ExchangeMailUtil(String host, String user, String password,String to) {
        super();
//        this.mailServer = "https://"+host+"/EWS/exchange.asmx";
        this.mailServer = "https://"+host+"/ecp";
        this.user = user;
        this.password = password;
        this.to = to;
    }
    /**
     * 创建邮件服务
     *
     * @return 邮件服务
     */
    private ExchangeService getExchangeService(){
        //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
        ExchangeService service = new ExchangeServiceWithHostVerify(ExchangeVersion.Exchange2010);
        // 用户认证信息
        ExchangeCredentials credentials = new WebCredentials(user, password);
        service.setCredentials(credentials);
        //代理
//        service.setWebProxy(new WebProxy("192.168.1.187",80));
        try {
            service.setUrl(new URI(mailServer));
        } catch (URISyntaxException e) {
            log.info("getExchangeService error {} ", e);
        }
        return service;
    }

    public void  receive (int max) throws Exception{
        ExchangeService service = getExchangeService();
        ExchangeCredentials credentials = new WebCredentials(to,password);
        service.setCredentials(credentials);
        service.setUrl(new URI(mailServer));
        //绑定收件箱,同样可以绑定发件箱
        Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);
        //获取文件总数量
        int count = inbox.getTotalCount();
        if (max > 0) {
            count = count > max ? max : count;
        }
        //循环获取邮箱邮件
        ItemView view = new ItemView(count);
        //按照时间顺序收取
        view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Descending);
        FindItemsResults<Item> findResults;
        findResults = service.findItems(inbox.getId(), view);
        service.loadPropertiesForItems(findResults, PropertySet.FirstClassProperties);
        for (Item item : findResults.getItems()) {
            EmailMessage message = (EmailMessage) item;
            //System.out.println(message.getSubject()+":"+message.getDateTimeReceived()+":"+message.getFrom().getName()+":"+message.getBody().toString()+":"+message.getIsRead());
            List<Attachment> attachmentList = message.getAttachments().getItems();
            for (Attachment attach : attachmentList) {
                if (attach instanceof FileAttachment) {
                    //接收邮件到临时目录
                    //System.out.println(attach.getName());
                    File tempZip = new File("d:/tmp", System.currentTimeMillis()+".jar");
                    ((FileAttachment) attach).load(tempZip.getPath());
                    tempZip.delete();
                }
            }
        }
    }

    /**
     * @param attachmentPath 附件
     * @throws Exception
     */
    public void send(String attachmentPath) throws Exception {
        ExchangeService service = getExchangeService();
        EmailMessage msg = new EmailMessage(service);
        String subject = "网关测试123";
        String bodyText = "对酒当歌，人生几何！\n" +
                "譬如朝露，去日苦多。\n" +
                "慨当以慷，忧思难忘。\n" +
                "何以解忧？唯有杜康。\n" +
                "青青子衿，悠悠我心。\n" +
                "但为君故，沉吟至今。\n" +
                "呦呦鹿鸣，食野之苹。\n" +
                "我有嘉宾，鼓瑟吹笙。\n" +
                "明明如月，何时可掇？\n" +
                "忧从中来，不可断绝。\n" +
                "越陌度阡，枉用相存。\n" +
                "契阔谈䜩，心念旧恩。\n" +
                "月明星稀，乌鹊南飞。\n" +
                "绕树三匝，何枝可依？\n" +
                "山不厌高，海不厌深。\n" +
                "周公吐哺，天下归心。";
        msg.setSubject(subject);
        MessageBody body = MessageBody.getMessageBodyFromText(bodyText);
        body.setBodyType(BodyType.HTML);
        msg.setBody(body);
        msg.getAttachments().addFileAttachment(attachmentPath);
        msg.getToRecipients().add(to);
        msg.send();
    }


    public static void main(String[] args) throws Exception  {
        String host = "192.168.1.83";
        String from = "ptest1@autotest.com";
        String pwd = "123qwe!@#";
        String to = "ptest2@autotest.com";
        // String url = "https://"+host+"/EWS/exchange.asmx";
        String attachmentPath = "D:\\apache-jmeter-3.2\\lib\\bsf-2.4.0.jar";
        ExchangeMailUtil mailUtil =
                new ExchangeMailUtil(host,from, pwd,to);
        //mailUtil.send(attachmentPath);
        mailUtil.receive(10);
    }
}
