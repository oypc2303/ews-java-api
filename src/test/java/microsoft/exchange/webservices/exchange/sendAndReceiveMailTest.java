package microsoft.exchange.webservices.exchange;

import lombok.extern.slf4j.Slf4j;
import microsoft.exchange.webservices.Exchange;
import microsoft.exchange.webservices.ExchangeMail;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import org.junit.Test;

import java.io.File;
import java.net.URI;
import java.util.List;

/**
 * @author oypc2
 * @version 1.0.0
 * @title email
 * @date 2022年10月12日 10:28:34
 * @description TODO
 */
@Slf4j
public class sendAndReceiveMailTest {
    String host = "192.168.201.115";
    String serverName = "https://192.168.201.115/ecp";
    String username = "mytest.com/Administrator";
    String password = "Holystone123";
    String from = "Administrator@mytest.com";
    String to = "Administrator@mytest.com";
    //附件路径
    String attachmentPath = "D:\\test.txt";

    @Test
    public void sendMail() throws Exception {
        ExchangeMailUtil exchangeMailUtil = new ExchangeMailUtil(host,from,password,to);
        exchangeMailUtil.send(attachmentPath);
    }

    @Test
    public void receiveMail() throws Exception {
        ExchangeMailUtil exchangeMailUtil = new ExchangeMailUtil(host,username,password,to);
        exchangeMailUtil.receive(100);

    }
    /**
     * java如何接收邮件_java Exchange服务接收邮件
     * https://blog.csdn.net/weixin_29960041/article/details/114653996
     * @param serverName
     * @param user
     * @param pwd
     * @param path
     * @param max
     * @return
     * @throws Exception
     */
    public int receive(String serverName, String user, String pwd, String path, int max) throws Exception {
        //新建ExchangeVersion.Exchange2007_SP1版本的Exchange服务
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
        String[] userInfo = user.split("/");
        //用户认证信息
        ExchangeCredentials credentials = new WebCredentials(userInfo[1], pwd, userInfo[0]);
        service.setCredentials(credentials);
        //设置Exchange连接的服务器地址
        service.setUrl(new URI(serverName));
        //绑定邮箱
        Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);
        //获取邮箱文件数量
        int count = inbox.getTotalCount();

        if(max > 0) {
            count = count > max ? max : count;
        }
        //循环获取邮箱邮件
        ItemView view = new ItemView(count);
        FindItemsResults<Item> findResults = service.findItems(inbox.getId(), view);
        for (Item item : findResults.getItems()) {
            EmailMessage message = EmailMessage.bind(service, item.getId());
            List<Attachment> attachs = message.getAttachments().getItems();
            try{
                if(message.getHasAttachments()){
                    for(Attachment f : attachs){
                        if(f instanceof FileAttachment){
                            //接收邮件到临时目录
                            File tempZip = new File(path, f.getName());
                            ((FileAttachment)f).load(tempZip.getPath());
                        }
                    }
                    //删除邮件
                    message.delete(DeleteMode.HardDelete);
                }
            }catch(Exception err){
                log.equals(err);
            }
        }
        return count;
    }
}
