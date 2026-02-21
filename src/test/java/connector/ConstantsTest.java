package connector;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link _Constants}.
 * Verifies that every constant holds the expected value so that accidental
 * edits to the shared configuration are caught immediately.
 */
class ConstantsTest {

    @Test
    void defaultProtocol() {
        assertEquals("http://", _Constants.DEFAULT_PROTOCOL);
    }

    @Test
    void defaultSslProtocol() {
        assertEquals("https://", _Constants.DEFAULT_SSL_PROTOCOL);
    }

    @Test
    void defaultLoginPath() {
        assertEquals("/_forms/default.aspx?wa=wsignin1.0", _Constants.DEFAULT_LOGIN_PATH);
    }

    @Test
    void defaultListDataPath() {
        assertEquals("/_vti_bin/ListData.svc/", _Constants.DEFAULT_LIST_DATA_PATH);
    }

    @Test
    void wsdlPath() {
        assertEquals("/_vti_bin/Lists.asmx?WSDL", _Constants.WSDL_PATH);
    }

    @Test
    void soapUrl() {
        assertEquals("http://schemas.microsoft.com/sharepoint/soap/", _Constants.SOAP_URL);
    }

    @Test
    void extstsUrl() {
        assertEquals("https://login.microsoftonline.com/extSTS.srf", _Constants.EXTSTS_SRF_URL);
    }

    @Test
    void utf8Charset() {
        assertEquals("UTF-8", _Constants.UTF_8_CHARSET);
    }

    @Test
    void defaultUserAgentWindowsIsNotBlank() {
        assertFalse(_Constants.DEFAULT_USER_AGENT_WINDOWS.isBlank());
    }

    @Test
    void defaultUserAgentMacIsNotBlank() {
        assertFalse(_Constants.DEFAULT_USER_AGENT_MAC.isBlank());
    }
}
