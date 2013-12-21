package connector;

import java.util.HashMap;
import java.util.Map;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

/**
 * @author David Dudok de Wit
 * @version 13 Feb. 2010
 * @see http://davidsit.wordpress.com/tag/connector/
 */
public class ListsRequest {

    private Document rootDocument;
    private Element rootDocContent;

    /**
     * @return the rootDocument
     */
    public Document getRootDocument() {
        return rootDocument;
    }

    /**
     * @return the rootDocContent
     */
    public Element getRootDocContent() {
        return rootDocContent;
    }

    /**
     * This class creates a generic XML SOAP request pre-formatted for
     * SharePoint Lists web services requests (aka CAML query). What remains to
     * be added are the specific parameters (XML Elements with attributes). For
     * an example of a CAML Doc
     * http://msdn.microsoft.com/en-us/library/lists.lists.updatelistitems.aspx
     *
     * @param requestType Either New, Update or Delete
     * @throws Exception
     */
    public ListsRequest(String requestType) throws Exception {
        if (requestType != null) {
            if (requestType.equals("New") || requestType.equals("Update") || requestType.equals("Delete")) {
                try {
                    Element rootElement; //= null;
                    DocumentBuilder docBuilder; //= null;
                    DocumentBuilderFactory dbfac = DocumentBuilderFactory.newInstance();
                    docBuilder = dbfac.newDocumentBuilder();
                    rootDocument = docBuilder.newDocument();

                    //Creates the root element
                    rootElement = rootDocument.createElement("Batch");
                    rootDocument.appendChild(rootElement);

                    //Creates the batch attributes
                    rootElement.setAttribute("ListVersion", "1");
                    rootElement.setAttribute("OnError", "Continue");
                    rootDocContent = rootDocument.createElement("Method");
                    rootDocContent.setAttribute("Cmd", requestType);
                    rootDocContent.setAttribute("ID", "1");
                    rootDocument.getElementsByTagName("Batch").item(0).appendChild(rootDocContent);
                } catch (ParserConfigurationException ex) {
                    //ex.printStackTrace();
                    throw (new Exception(ex.toString()));
                }
            } else {
                String err = "Unsupported request type";
                throw (new Exception(err));
            }
        } else {
            String err = "Null parameters";
            throw (new Exception(err));
        }
    }

    /**
     * Creates a SharePoint list item in the CAML format, and adds it to the
     * rootRequest. In SharePoint, this corresponds to a line in a list. The
     * parameters given here would correspond respectively to the name of the
     * column where to insert the info, and then the info itself. The
     * requestTypeElement should already be initialized before calling this
     * method. XML example output: < Field Name="ID" >4< Field > < Field
     * Name="Field_Name" >Value< /Field >
     *
     * @param fields Contains a HashMap with attribute names as keys, and
     * attributes values as content
     * @return true if the item has been successfully added to the caml request
     */
    public boolean createListItem(HashMap<String, String> fields) {
        //params check
        if (fields != null && getRootDocContent() != null && this.getRootDocument() != null && !fields.isEmpty()) {
            Element createdElement; // = null;
            //Adds attribute by attribute to fields
            for (Map.Entry<String, String> aField : fields.entrySet()) {
                createdElement = this.getRootDocument().createElement("Field");
                createdElement.setAttribute("Name", aField.getKey());
                Text attributeValue = getRootDocument().createTextNode("" + aField.getValue());
                createdElement.appendChild(attributeValue);
                this.getRootDocContent().appendChild(createdElement);
            }
            return true;
        }
        return false;
    }
}
