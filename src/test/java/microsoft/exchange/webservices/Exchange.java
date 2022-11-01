package microsoft.exchange.webservices;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Vector;
/**
 * @author oypc2
 * @version 1.0.0
 * @title Exchange
 * @date 2022年10月12日 10:23:07
 * @description TODO
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class Exchange {
    private String displayName;
    private String to;
    private String from;
    private String smtpServer;
    private String username;
    private String password;
    private String subject;
    private String content;
    private boolean isExchange;
    private String domain;
    private boolean ifAuth;
    private String filename;
    private Vector<String> file = new Vector<String>(); // 用于保存发送附件的文件名的集合

    /**
     * @param smtpServer 发送服务器地址
     * @param from 发送人地址
     * @param displayName 发送人发送名
     * @param to 接收人
     * @param subject 主题
     * @param content 内容
     * @param isExchange 是否验证
     * @param domain 域名
     * @param username 用户登录名
     * @param password 密码
     */
    public Exchange(String smtpServer, String from, String displayName, String to,
                    String subject, String content, boolean isExchange, String domain,
                    String username, String password) {
        this.smtpServer = smtpServer;
        this.from = from;
        this.displayName = displayName;
        this.ifAuth = true;
        this.to = to;
        this.subject = subject;
        this.content = content;
        this.isExchange = isExchange;
        this.domain = domain;
        this.username = username;
        this.password = password;
    }
}