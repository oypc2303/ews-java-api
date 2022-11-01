package microsoft.exchange.webservices.data;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.ServletComponentScan;
import org.springframework.scheduling.annotation.EnableScheduling;

/**
 * @author oypc2
 * @version 1.0.0
 * @title EwsApiApplication
 * @date 2022年10月17日 16:17:39
 * @description TODO
 */
@EnableScheduling
@SpringBootApplication
@ServletComponentScan
public class EwsApiApplication {
    public static void main(String[] args) {
        SpringApplication.run(EwsApiApplication.class, args);
    }
}
