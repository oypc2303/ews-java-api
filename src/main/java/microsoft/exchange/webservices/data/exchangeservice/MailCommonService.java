package microsoft.exchange.webservices.data.exchangeservice;


import lombok.extern.slf4j.Slf4j;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import org.apache.commons.lang3.StringUtils;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Component;
import org.springframework.util.CollectionUtils;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Address;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.*;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URI;
import java.util.*;

/**
 * @author oypc2
 * @version 1.0.0
 * @title MailCommonService
 * @date 2022年10月17日 15:44:39
 * @description TODO
 */
@Component
public class MailCommonService {
    /**
     *
     * @param toList 收件人列表
     * @param mailSubject 邮件主题
     * @param contentList 邮件正文，会按照列表的顺序，在邮件中显示。
     *                    邮件内容有多行时，每行内容是一个MailContentVO
     * @param fileList 附件文件绝对路径
     * @param type 邮件协议类型
     */
    public void sendMail(List<String> toList,String mailSubject,List<MailContentVO> contentList,List<String> fileList,MailProtocolType type) {
        try {
            Properties props = new Properties();
            ClassPathResource classPathResource = new ClassPathResource("/mail.properties");
            InputStream inputStream  = classPathResource.getInputStream();
            props.load(new InputStreamReader(inputStream, "UTF-8"));
            if (type.equals(MailProtocolType.smtp)){
                sendMailSmtp(toList,mailSubject,contentList,fileList,props);
            } else if(type.equals(MailProtocolType.exchange)){
                sendMailExchange(toList,mailSubject,contentList,fileList,props);
            } else {
                sendMailSmtp(toList,mailSubject,contentList,fileList,props);
            }
            System.out.println("邮件发送完成......");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 发送邮件
     * @param toList 收件人列表
     * @param mailSubject 邮件主题
     * @param contentList 邮件正文，会按照列表的顺序，在邮件中显示。
     *                    邮件内容有多行时，每行内容是一个MailContentVO
     * @param fileList 附件文件绝对路径
     * @param properties 发送邮箱相关配置
     */
    public void sendMailSmtp(List<String> toList, String mailSubject, List<MailContentVO> contentList, List<String> fileList, Properties properties) {
        String senderAddress = properties.getProperty("smtp.send.address");
        String senderPassword = properties.getProperty("smtp.send.password");
        try {
            //1、连接邮件服务器的参数配置
            Properties props = new Properties();
            //设置用户的认证方式
            props.setProperty("mail.smtp.auth", "true");
            //设置传输协议
            props.setProperty("mail.transport.protocol", "smtp");
            //设置发件人的SMTP服务器地址
            props.setProperty("mail.smtp.host", properties.getProperty("smtp.send.host"));
            //2、创建定义整个应用程序所需的环境信息的 Session 对象
            Session session = Session.getInstance(props);
            //设置调试信息在控制台打印出来
            session.setDebug(true);
            //3、创建邮件的实例对象
            Message msg = getMimeMessage(session,senderAddress,toList,mailSubject,contentList,fileList);
            //4、根据session对象获取邮件传输对象Transport
            Transport transport = session.getTransport();
            //设置发件人的账户名和密码
            transport.connect(senderAddress, senderPassword);
            //发送邮件,如果只想发送给指定的人，可以如下写法
            Address[] addresses = new Address[toList.size()];
            for (int i=0;i<toList.size();i++){
                addresses[i] = new InternetAddress(toList.get(i));
            }
            transport.sendMessage(msg, addresses);
            //5、关闭邮件连接
            transport.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private MimeMessage getMimeMessage(Session session, String senderAddress, List<String> toList, String mailSubject, List<MailContentVO> contentList, List<String> fileList) throws Exception {
        //1、创建一封邮件的实例对象
        MimeMessage msg = new MimeMessage(session);
        //2、设置发件人地址
        msg.setFrom(new InternetAddress(senderAddress));
        /**
         * 3、设置收件人地址（可以增加多个收件人、抄送、密送），即下面这一行代码书写多行
         * MimeMessage.RecipientType.TO:发送
         * MimeMessage.RecipientType.CC：抄送
         * MimeMessage.RecipientType.BCC：密送
         */
        for (String to:toList){
            msg.setRecipient(MimeMessage.RecipientType.TO, new InternetAddress(to));
        }
        //4、设置邮件主题
        msg.setSubject(mailSubject, "UTF-8");

        // 5、（文本+图片）设置 文本 和 图片"节点"的关系（将 文本 和 图片"节点"合成一个混合"节点"）
        MimeMultipart mm_text_image = new MimeMultipart();

        //文本节点的内容组合
        List<String> textContent = new ArrayList<String>();
        for (MailContentVO vo:contentList){
            MailContentType type = vo.getType();
            if (type.equals(MailContentType.txt)){
                textContent.add(vo.getContent());
            }else{
                String uid = UUID.randomUUID().toString();
                //创建图片"节点"
                MimeBodyPart image = new MimeBodyPart();
                //读取本地文件
                DataHandler dh = new DataHandler(new FileDataSource(vo.getContent()));
                //将图片数据添加到"节点"
                image.setDataHandler(dh);
                //为"节点"设置一个唯一编号（在文本"节点"将引用该ID）
                image.setContentID(uid);
                mm_text_image.addBodyPart(image);
                //
                String img = "<img src='cid:"+uid+"'/>";
                textContent.add(img);
            }
        }

        MimeBodyPart text = new MimeBodyPart();
        // 这里添加图片的方式是将整个图片包含到邮件内容中, 实际上也可以以 http 链接的形式添加网络图片
        text.setContent(StringUtils.join(textContent,"<br/>"), "text/html;charset=UTF-8");
        mm_text_image.addBodyPart(text);
        mm_text_image.setSubType("related"); //关联关系

        // 8. 将 文本+图片 的混合"节点"封装成一个普通"节点"
        // 最终添加到邮件的 Content 是由多个 BodyPart 组成的 Multipart, 所以我们需要的是 BodyPart,
        MimeBodyPart text_image = new MimeBodyPart();
        text_image.setContent(mm_text_image);

        MimeMultipart mm = new MimeMultipart();
        mm.addBodyPart(text_image);
        //9、设置附件
        for (String filePath:fileList){
            File file = new File(filePath);
            MimeBodyPart attachment = new MimeBodyPart();
            DataSource source = new FileDataSource(filePath);
            attachment.setDataHandler(new DataHandler(source));
            attachment.setFileName(MimeUtility.encodeText(file.getName()));
            mm.addBodyPart(attachment);
        }
        mm.setSubType("mixed");  // 混合关系
        // 10、 设置整个邮件的关系（将最终的混合"节点"作为邮件的内容添加到邮件对象）
        msg.setContent(mm);
        //设置邮件的发送时间,默认立即发送
        msg.setSentDate(new Date());
        return msg;
    }

    /**
     * 发送邮件
     * @param toList 收件人列表
     * @param subject 邮件主题
     * @param contentList 邮件正文，会按照列表的顺序，在邮件中显示。
     *                    邮件内容有多行时，每行内容是一个MailContentVO
     * @param filePath 附件文件绝对路径
     * @param properties 发送邮箱相关配置
     */
    private void sendMailExchange(List<String> toList, String subject, List<MailContentVO> contentList, List<String> filePath, Properties properties)  {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
        ExchangeCredentials credentials = new WebCredentials(properties.getProperty("exchange.send.user"),properties.getProperty("exchange.send.password"));
        service.setCredentials(credentials);
        try {
            String uri = "https://"+properties.getProperty("exchange.send.host")+"/ews/exchange.asmx";
            service.setUrl(new URI(uri));
            EmailMessage msg = new EmailMessage(service);
            //设置邮件主题
            msg.setSubject(subject);

            //设置邮件正文
            List<String> textContent = new ArrayList<String>();
            for (MailContentVO vo:contentList){
                MailContentType type = vo.getType();
                if (type.equals(MailContentType.txt)){
                    textContent.add(vo.getContent());
                }else{
                    String uid = UUID.randomUUID().toString();
                    FileAttachment attachment = msg.getAttachments().addFileAttachment(vo.getContent());
                    attachment.setContentType("image");
                    attachment.setContentId(uid);
                    String img = "<img src='cid:"+uid+"'/>";
                    textContent.add(img);
                }
            }

            //设置邮件附件
            if (!CollectionUtils.isEmpty(filePath)){
                for (String path:filePath){
                    String uid = UUID.randomUUID().toString();
                    FileAttachment attachment1 = msg.getAttachments().addFileAttachment(path);
                    attachment1.setContentType("image");
                    attachment1.setContentId(uid);
                }
            }
            MessageBody body = MessageBody.getMessageBodyFromText(StringUtils.join(textContent,"<br/>")+"<hr/>");
            body.setBodyType(BodyType.HTML);
            msg.setBody(body);

            //设置邮件收件人
            for (String to : toList) {
                msg.getToRecipients().add(to);
            }

            //邮件发送
            msg.send();
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
}
