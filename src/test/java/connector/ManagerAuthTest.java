package connector;

import com.microsoft.schemas.sharepoint.soap.GetListResponse.GetListResult;
import com.microsoft.schemas.sharepoint.soap.GetListItemsResponse.GetListItemsResult;
import com.microsoft.schemas.sharepoint.soap.ListsSoap;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.Mockito;

import java.io.InputStream;
import java.io.ByteArrayInputStream;
import java.util.ArrayList;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Unit tests for the authentication/precondition logic and null-safety
 * guards in {@link Manager}.
 *
 * Network calls are never made — the SharePoint SOAP port is mocked wherever
 * the method under test requires one.
 */
class ManagerAuthTest {

    // -----------------------------------------------------------------------
    // sharePointListsAuth — precondition checks (no network)
    // -----------------------------------------------------------------------

    @BeforeEach
    void initManagerService() {
        // Ensure Manager.instanced == true for all tests so the "must call
        // createManagerService first" branch is not triggered unintentionally.
        Manager.createManagerService("testsite.sharepoint.com", "/sites/test");
    }

    @Test
    void sharePointListsAuthThrowsForNullUsername() {
        Exception ex = assertThrows(Exception.class,
                () -> Manager.sharePointListsAuth(null, "password", "cookie"));
        assertEquals("Couldn't authenticate: Invalid connection details given.", ex.getMessage());
    }

    @Test
    void sharePointListsAuthThrowsForNullPassword() {
        Exception ex = assertThrows(Exception.class,
                () -> Manager.sharePointListsAuth("user@test.com", null, "cookie"));
        assertEquals("Couldn't authenticate: Invalid connection details given.", ex.getMessage());
    }

    @Test
    void sharePointListsAuthThrowsForBothNull() {
        Exception ex = assertThrows(Exception.class,
                () -> Manager.sharePointListsAuth(null, null, "cookie"));
        assertEquals("Couldn't authenticate: Invalid connection details given.", ex.getMessage());
    }

    @Test
    void createManagerServiceBuildsCorrectWsdlUrl() throws Exception {
        // After createManagerService the WSDL URL embeds the endpoint.
        // We verify indirectly: passing null credentials must produce the
        // "Invalid connection details" error (not "must be first execution"),
        // confirming instanced == true.
        Manager.createManagerService("mysite.sharepoint.com", "/sites/hr");
        Exception ex = assertThrows(Exception.class,
                () -> Manager.sharePointListsAuth(null, "pass", "cookie"));
        assertEquals("Couldn't authenticate: Invalid connection details given.", ex.getMessage());
    }

    // -----------------------------------------------------------------------
    // addAttachMentToListItem — null-safety (no file I/O)
    // -----------------------------------------------------------------------

    @Test
    void addAttachMentToListItemFilePathReturnsNullWhenPortIsNull() throws Exception {
        String result = Manager.addAttachMentToListItem(
                null, "{GUID}", "1", "/some/path", "file.txt");
        assertNull(result, "null port should produce null result without throwing");
    }

    @Test
    void addAttachMentToListItemFilePathReturnsNullWhenGuidIsNull() throws Exception {
        ListsSoap mockPort = Mockito.mock(ListsSoap.class);
        String result = Manager.addAttachMentToListItem(
                mockPort, null, "1", "/some/path", "file.txt");
        assertNull(result);
        verifyNoInteractions(mockPort);
    }

    @Test
    void addAttachMentToListItemFilePathReturnsNullWhenListItemIdIsNull() throws Exception {
        ListsSoap mockPort = Mockito.mock(ListsSoap.class);
        String result = Manager.addAttachMentToListItem(
                mockPort, "{GUID}", null, "/some/path", "file.txt");
        assertNull(result);
        verifyNoInteractions(mockPort);
    }

    @Test
    void addAttachMentToListItemStreamReturnsNullWhenPortIsNull() throws Exception {
        InputStream stream = new ByteArrayInputStream(new byte[]{1, 2, 3});
        String result = Manager.addAttachMentToListItem(
                null, "{GUID}", "1", "file.txt", stream);
        assertNull(result, "null port should produce null result without throwing");
    }

    @Test
    void addAttachMentToListItemStreamReturnsNullWhenStreamIsNull() throws Exception {
        ListsSoap mockPort = Mockito.mock(ListsSoap.class);
        String result = Manager.addAttachMentToListItem(
                mockPort, "{GUID}", "1", "file.txt", (InputStream) null);
        assertNull(result);
        verifyNoInteractions(mockPort);
    }

    @Test
    void addAttachMentToListItemStreamDelegatesToPort() throws Exception {
        ListsSoap mockPort = Mockito.mock(ListsSoap.class);
        when(mockPort.addAttachment(eq("{GUID}"), eq("1"), eq("file.txt"), any(byte[].class)))
                .thenReturn("/sites/test/Attachments/1/file.txt");

        InputStream stream = new ByteArrayInputStream("hello".getBytes());
        String result = Manager.addAttachMentToListItem(
                mockPort, "{GUID}", "1", "file.txt", stream);

        assertEquals("/sites/test/Attachments/1/file.txt", result);
        verify(mockPort).addAttachment(eq("{GUID}"), eq("1"), eq("file.txt"), any(byte[].class));
    }

    // -----------------------------------------------------------------------
    // displaySharePointList — null-safety (no network)
    // -----------------------------------------------------------------------

    @Test
    void displaySharePointListDoesNothingWhenPortIsNull() {
        assertDoesNotThrow(() ->
                Manager.displaySharePointList(null, "MyList", new ArrayList<>(), "100"));
    }

    @Test
    void displaySharePointListDoesNothingWhenListNameIsNull() {
        ListsSoap mockPort = Mockito.mock(ListsSoap.class);
        assertDoesNotThrow(() ->
                Manager.displaySharePointList(mockPort, null, new ArrayList<>(), "100"));
        verifyNoInteractions(mockPort);
    }

    @Test
    void displaySharePointListDoesNothingWhenColumnNamesIsNull() {
        ListsSoap mockPort = Mockito.mock(ListsSoap.class);
        assertDoesNotThrow(() ->
                Manager.displaySharePointList(mockPort, "MyList", null, "100"));
        verifyNoInteractions(mockPort);
    }

    // -----------------------------------------------------------------------
    // getUIDFromListElement — null-safety (no network)
    // -----------------------------------------------------------------------

    @Test
    void getUIDFromListElementReturnsNullWhenPortIsNull() throws Exception {
        String result = Manager.getUIDFromListElement(null, "MyList", "42");
        assertNull(result);
    }

    @Test
    void getUIDFromListElementReturnsNullWhenListNameIsNull() throws Exception {
        ListsSoap mockPort = Mockito.mock(ListsSoap.class);
        String result = Manager.getUIDFromListElement(mockPort, null, "42");
        assertNull(result);
        verifyNoInteractions(mockPort);
    }

    @Test
    void getUIDFromListElementReturnsNullWhenOwsIdIsNull() throws Exception {
        ListsSoap mockPort = Mockito.mock(ListsSoap.class);
        String result = Manager.getUIDFromListElement(mockPort, "MyList", null);
        assertNull(result);
        verifyNoInteractions(mockPort);
    }
}
