package microsoft.exchange.webservices.exchange;
import microsoft.exchange.webservices.data.EWSConstants;
import microsoft.exchange.webservices.data.core.EwsSSLProtocolSocketFactory;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import org.apache.http.config.Registry;
import org.apache.http.config.RegistryBuilder;
import org.apache.http.conn.socket.ConnectionSocketFactory;
import org.apache.http.conn.socket.PlainConnectionSocketFactory;
import javax.net.ssl.HostnameVerifier;
import javax.net.ssl.SSLSession;
import java.security.GeneralSecurityException;

/**
 * @author oypc2
 * @version 1.0.0
 * @title ExchangeServiceWithHostVerify
 * @date 2022年10月14日 13:36:35
 * @description TODO
 */
public class ExchangeServiceWithHostVerify extends ExchangeService {
    private final static HostnameVerifier hostnameVerifierWithOutVerfy = new HostnameVerifier() {

        public boolean verify(String s, SSLSession sslSession) {
            return true;
        }
    };


    public ExchangeServiceWithHostVerify(ExchangeVersion requestedServerVersion) {
        super(requestedServerVersion);
    }

    protected Registry<ConnectionSocketFactory> createConnectionSocketFactoryRegistry() {
        try {
            return RegistryBuilder.<ConnectionSocketFactory>create()
                    .register(EWSConstants.HTTP_SCHEME, new PlainConnectionSocketFactory())
                    .register(EWSConstants.HTTPS_SCHEME,
                            EwsSSLProtocolSocketFactory.build(null, hostnameVerifierWithOutVerfy))
                    .build();
        } catch (GeneralSecurityException e) {
            throw new RuntimeException(
                    "Could not initialize ConnectionSocketFactory instances for HttpClientConnectionManager", e);
        }
    }
}
