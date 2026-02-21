package connector;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.HashMap;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ListsRequest}.
 *
 * Covers CAML batch XML structure construction, the three supported request
 * types, field element generation, and all invalid-input paths.
 */
class ListsRequestTest {

    // -----------------------------------------------------------------------
    // Constructor — valid request types
    // -----------------------------------------------------------------------

    @Test
    void newRequestCreatesNonNullDocumentAndMethod() throws Exception {
        ListsRequest req = new ListsRequest("New");
        assertNotNull(req.getRootDocument());
        assertNotNull(req.getRootDocContent());
    }

    @Test
    void newRequestCreatesBatchRootElement() throws Exception {
        ListsRequest req = new ListsRequest("New");
        Document doc = req.getRootDocument();

        NodeList batches = doc.getElementsByTagName("Batch");
        assertEquals(1, batches.getLength(), "Exactly one <Batch> element expected");

        Element batch = (Element) batches.item(0);
        assertEquals("1", batch.getAttribute("ListVersion"));
        assertEquals("Continue", batch.getAttribute("OnError"));
    }

    @Test
    void newRequestSetsMethodCmdNew() throws Exception {
        ListsRequest req = new ListsRequest("New");
        Element method = req.getRootDocContent();
        assertEquals("New", method.getAttribute("Cmd"));
        assertEquals("1", method.getAttribute("ID"));
    }

    @Test
    void updateRequestSetsMethodCmdUpdate() throws Exception {
        ListsRequest req = new ListsRequest("Update");
        assertEquals("Update", req.getRootDocContent().getAttribute("Cmd"));
    }

    @Test
    void deleteRequestSetsMethodCmdDelete() throws Exception {
        ListsRequest req = new ListsRequest("Delete");
        assertEquals("Delete", req.getRootDocContent().getAttribute("Cmd"));
    }

    // -----------------------------------------------------------------------
    // Constructor — invalid input
    // -----------------------------------------------------------------------

    @Test
    void unsupportedTypeThrowsWithDescriptiveMessage() {
        Exception ex = assertThrows(Exception.class, () -> new ListsRequest("Read"));
        assertEquals("Unsupported request type", ex.getMessage());
    }

    @Test
    void emptyStringTypeThrowsUnsupportedRequestType() {
        Exception ex = assertThrows(Exception.class, () -> new ListsRequest(""));
        assertEquals("Unsupported request type", ex.getMessage());
    }

    @Test
    void nullTypeThrowsWithDescriptiveMessage() {
        Exception ex = assertThrows(Exception.class, () -> new ListsRequest(null));
        assertEquals("Null parameters", ex.getMessage());
    }

    // -----------------------------------------------------------------------
    // createListItem — happy path
    // -----------------------------------------------------------------------

    @Test
    void createListItemReturnsTrueForValidFields() throws Exception {
        ListsRequest req = new ListsRequest("New");
        HashMap<String, String> fields = new HashMap<>();
        fields.put("Title", "Hello World");

        assertTrue(req.createListItem(fields));
    }

    @Test
    void createListItemAddsOneFieldElementPerEntry() throws Exception {
        ListsRequest req = new ListsRequest("New");
        HashMap<String, String> fields = new HashMap<>();
        fields.put("Title", "Test Title");
        fields.put("Description", "Test Desc");

        req.createListItem(fields);

        NodeList fieldNodes = req.getRootDocument().getElementsByTagName("Field");
        assertEquals(2, fieldNodes.getLength());
    }

    @Test
    void createListItemSetsNameAttributeAndTextContent() throws Exception {
        ListsRequest req = new ListsRequest("New");
        HashMap<String, String> fields = new HashMap<>();
        fields.put("Title", "My Value");

        req.createListItem(fields);

        NodeList fieldNodes = req.getRootDocument().getElementsByTagName("Field");
        assertEquals(1, fieldNodes.getLength());
        Element field = (Element) fieldNodes.item(0);
        assertEquals("Title", field.getAttribute("Name"));
        assertEquals("My Value", field.getTextContent());
    }

    @Test
    void createListItemIsIdempotentAndAccumulatesFields() throws Exception {
        ListsRequest req = new ListsRequest("New");
        HashMap<String, String> first = new HashMap<>();
        first.put("Title", "First");
        HashMap<String, String> second = new HashMap<>();
        second.put("Body", "Second");

        req.createListItem(first);
        req.createListItem(second);

        NodeList fieldNodes = req.getRootDocument().getElementsByTagName("Field");
        assertEquals(2, fieldNodes.getLength());
    }

    // -----------------------------------------------------------------------
    // createListItem — invalid input
    // -----------------------------------------------------------------------

    @Test
    void createListItemReturnsFalseForNullFields() throws Exception {
        ListsRequest req = new ListsRequest("New");
        assertFalse(req.createListItem(null));
    }

    @Test
    void createListItemReturnsFalseForEmptyFields() throws Exception {
        ListsRequest req = new ListsRequest("New");
        assertFalse(req.createListItem(new HashMap<>()));
    }
}
