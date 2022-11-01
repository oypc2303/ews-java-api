package microsoft.exchange.webservices;

import lombok.extern.slf4j.Slf4j;
import microsoft.exchange.webservices.data.EwsApiApplication;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.exchangeservice.MailCommonService;
import microsoft.exchange.webservices.data.exchangeservice.MailContentType;
import microsoft.exchange.webservices.data.exchangeservice.MailContentVO;
import microsoft.exchange.webservices.data.exchangeservice.MailProtocolType;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;
import microsoft.exchange.webservices.exchange.ExchangeServiceWithHostVerify;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.net.URI;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * @author oypc2
 * @version 1.0.0
 * @title email
 * @date 2022年10月12日 10:28:34
 * @description TODO
 */
@Slf4j
@RunWith(SpringRunner.class)
@SpringBootTest(classes = EwsApiApplication.class)
public class testExchangeMail {
    @Autowired
    MailCommonService mailCommonService;

    @Before
    @Test
    public void testSendMail(){
        List<String> toList = new ArrayList<String>();
        toList.add("ouyang@mytest1.com");
        List<String> fileList = new ArrayList<String>();
        fileList.add("D:\\upload\\test2.txt");
        List<MailContentVO> mailContentVOList = new ArrayList<MailContentVO>();
        MailContentVO mailContentVO = new MailContentVO();
        mailContentVO.setContent("qasd");
        mailContentVO.setType(MailContentType.txt);
        mailContentVOList.add(mailContentVO);

//        mailCommonService.sendMail(toList,"测试",
//                mailContentVOList,fileList, MailProtocolType.exchange);
        mailCommonService.sendMail(toList,"测试",
                mailContentVOList,fileList, MailProtocolType.smtp);
    }


    @Test
    public void receiveMail() throws Exception {
        String serverName = "https://mail.mytest1.com/EWS/Exchange.asmx";
        String username = "mytest1.com/ouyangpengcheng";
        String password = "Holystone123";
        String path = "D:/upload/1/";
        int max = -1;
        receive(serverName, username, password, path, max);
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
        ExchangeService service = new ExchangeServiceWithHostVerify(ExchangeVersion.Exchange2007_SP1);
        String[] userInfo = user.split("/");
        //用户认证信息
        ExchangeCredentials credentials = new WebCredentials(userInfo[1], pwd, userInfo[0]);
        service.setUseDefaultCredentials(true);
        service.setCredentials(credentials);
        //设置Exchange连接的服务器地址
        service.setUrl(new URI(serverName));
//        service.autodiscoverUrl("oypc2303@mytest.com");
        //绑定邮箱
        Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);
        //获取邮箱文件数量
        int count = inbox.getTotalCount();

        if(max > 0) {
            count = count > max ? max : count;
        }
        //循环获取邮箱邮件
        ItemView view = new ItemView(count);
        FindItemsResults<Item> findResults = service.findItems(inbox.getId(),
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false), view);
        for (Item item : findResults.getItems()) {
            EmailMessage message = EmailMessage.bind(service, item.getId());
            //https://blog.csdn.net/Sayonara_LM/article/details/107711147
            System.out.println("邮件内容：" + message.getBody().toString());
            String text = message.getBody().toString();
            System.out.println("邮件内容：" + Html2Text(text));
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
//                    message.delete(DeleteMode.HardDelete);
                }
            }catch(Exception err){
                log.equals(err);
            }
            //更新为已读
            message.update(ConflictResolutionMode.AlwaysOverwrite);
        }
        return count;
    }

    public static String Html2Text(String inputString) {
        String htmlStr = inputString; // 含html标签的字符串
        String textStr = "";
        java.util.regex.Pattern p_script;
        java.util.regex.Matcher m_script;
        java.util.regex.Pattern p_style;
        java.util.regex.Matcher m_style;
        java.util.regex.Pattern p_html;
        java.util.regex.Matcher m_html;
        try {
            String regEx_script = "<[\\s]*?script[^>]*?>[\\s\\S]*?<[\\s]*?\\/[\\s]*?script[\\s]*?>"; // 定义script的正则表达式{或<script[^>]*?>[\\s\\S]*?<\\/script>
            String regEx_style = "<[\\s]*?style[^>]*?>[\\s\\S]*?<[\\s]*?\\/[\\s]*?style[\\s]*?>"; // 定义style的正则表达式{或<style[^>]*?>[\\s\\S]*?<\\/style>
            String regEx_html = "<[^>]+>"; // 定义HTML标签的正则表达式
            p_script = Pattern.compile(regEx_script, Pattern.CASE_INSENSITIVE);
            m_script = p_script.matcher(htmlStr);
            htmlStr = m_script.replaceAll(""); // 过滤script标签
            p_style = Pattern.compile(regEx_style, Pattern.CASE_INSENSITIVE);
            m_style = p_style.matcher(htmlStr);
            htmlStr = m_style.replaceAll(""); // 过滤style标签
            p_html = Pattern.compile(regEx_html, Pattern.CASE_INSENSITIVE);
            m_html = p_html.matcher(htmlStr);
            htmlStr = m_html.replaceAll(""); // 过滤html标签
            textStr = htmlStr;
        } catch (Exception e) {System.err.println("Html2Text: " + e.getMessage()); }
        //剔除空格行
        textStr=textStr.replaceAll("[ ]+", " ");
        textStr=textStr.replaceAll("(?m)^\\s*$(\\n|\\r\\n)", "");
        //返回文本字符串
        return textStr;
    }
}
