package microsoft.exchange.webservices;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Properties;

/**
 * @author oypc2
 * @version 1.0.0
 * @title ExchangeMail
 * @date 2022年10月12日 10:27:33
 * @description TODO
 */
public class ExchangeMail {
    /**
     * 发送邮件
     */
    public static HashMap<String, String> send(Exchange exchange) {
        HashMap<String, String> map = new HashMap<String,String>();
        map.put("state", "success");
        String message = "邮件发送成功！";
        Session session = null;
        Properties props = System.getProperties();
        props.put("mail.smtp.host", exchange.getSmtpServer());
        if (exchange.isExchange()) {
            if (exchange.getDomain() == null || exchange.getDomain().equals("")) {
                throw new RuntimeException("domain is null");
            }
            props.setProperty("mail.smtp.auth.ntlm.domain", exchange.getDomain());
        }
        if (exchange.isIfAuth()) { // 服务器需要身份认证
            props.put("mail.smtp.auth", "true");
            MailAuthenticator smtpAuth = new MailAuthenticator(exchange.getUsername(), exchange.getPassword());
            session = Session.getDefaultInstance(props, smtpAuth);
        } else {
            props.put("mail.smtp.auth", "false");
            session = Session.getDefaultInstance(props, null);
        }
        session.setDebug(true);
        Transport trans = null;
        try {
            Message msg = new MimeMessage(session);
            try {
                Address from_address = new InternetAddress(exchange.getFrom(), exchange.getDisplayName());
                msg.setFrom(from_address);
            } catch (java.io.UnsupportedEncodingException e) {
                e.printStackTrace();
            }
            InternetAddress[] address = {new InternetAddress(exchange.getTo())};
            msg.setRecipients(Message.RecipientType.TO, address);
            msg.setSubject(exchange.getSubject());
            //设置传输文件
            Multipart mp = new MimeMultipart();
            MimeBodyPart mbp = new MimeBodyPart();
            mbp.setContent(exchange.getContent().toString(), "text/html;charset=gb2312");
            mp.addBodyPart(mbp);
            if (!exchange.getFile().isEmpty()) {// 有附件
                Enumeration<String> efile = exchange.getFile().elements();
                while (efile.hasMoreElements()) {
                    mbp = new MimeBodyPart();
                    exchange.setFilename(efile.nextElement().toString()); // 选择出每一个附件名
                    FileDataSource fds = new FileDataSource(exchange.getFilename()); // 得到数据源
                    mbp.setDataHandler(new DataHandler(fds)); // 得到附件本身并至入BodyPart
                    mbp.setFileName(fds.getName()); // 得到文件名同样至入BodyPart
                    mp.addBodyPart(mbp);
                }
                exchange.getFile().removeAllElements();
            }
            msg.setContent(mp); // Multipart加入到信件
            msg.setSentDate(new Date()); // 设置信件头的发送日期
            // 发送信件
            msg.saveChanges();
            trans = session.getTransport("smtp");
            trans.connect(exchange.getSmtpServer(), exchange.getUsername(), exchange.getPassword());
            trans.sendMessage(msg, msg.getAllRecipients());
            trans.close();
        } catch (AuthenticationFailedException e) {
            map.put("state", "failed");
            message = "邮件发送失败！错误原因：\n" + "身份验证错误!";
            e.printStackTrace();
        } catch (MessagingException e) {
            message = "邮件发送失败！错误原因：\n" + e.getMessage();
            map.put("state", "failed");
            e.printStackTrace();
            Exception ex = null;
            if ((ex = e.getNextException()) != null) {
                System.out.println(ex.toString());
                ex.printStackTrace();
            }
        }
        // System.out.println("\n提示信息:"+message);
        map.put("message", message);
        return map;
    }
}
