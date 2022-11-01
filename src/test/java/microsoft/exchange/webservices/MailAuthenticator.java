package microsoft.exchange.webservices;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.mail.Authenticator;
import javax.mail.PasswordAuthentication;

/**
 * @author oypc2
 * @version 1.0.0
 * @title MailAuthenticator
 * @date 2022年10月12日 10:26:19
 * @description TODO
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class MailAuthenticator extends Authenticator {
    /**
     * 用户名（登录邮箱）
     */

    private String username;
    /**
     * 密码
     */
    private String password;


    @Override
    protected PasswordAuthentication getPasswordAuthentication() {
        return new PasswordAuthentication(username, password);
    }
}