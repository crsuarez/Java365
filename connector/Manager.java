package connector;

import com.microsoft.schemas.sharepoint.soap.GetListResponse.GetListResult;
import com.microsoft.schemas.sharepoint.soap.*;
import com.microsoft.schemas.sharepoint.soap.UpdateListItems.Updates;
import com.microsoft.schemas.sharepoint.soap.UpdateListItemsResponse.UpdateListItemsResult;
import com.sun.org.apache.xerces.internal.dom.ElementNSImpl;
import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.namespace.QName;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.ws.BindingProvider;
import javax.xml.ws.handler.MessageContext;
import org.apache.commons.io.IOUtils;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

/**
 *
 * @author ingcarlos portions of code belongs to
 * https://davidsit.wordpress.com/2010/02/10/reading-a-sharepoint-list-with-java-tutorial/
 */
public class Manager {

    private static String _wsdlURL;
    private static final String _soapURL = _Constants.SOAP_URL;
    private static final Logger _out;
    private static boolean instanced = false;

    static {
        _out = Logger.getLogger(Manager.class.getName());
    }

    /**
     * This method do some assignation for the client.
     *
     * @param endPoint the url of your site, without the protocol. ie.
     * yoursite.sharepoint.com
     * @param relPathToListLocation the path in where is your list.
     */
    public static void createManagerService(String endPoint, String relPathToListLocation) {
        Manager._wsdlURL = _Constants.DEFAULT_SSL_PROTOCOL + endPoint + relPathToListLocation + _Constants.WSDL_PATH;
        Manager.instanced = true;
    }

    /**
     * Creates a port connected to the SharePoint Web Service given.
     * Authentication is done here. It also prints the authentication details in
     * the console.
     *
     * @param userName SharePoint username
     * @param password SharePoint password
     * @return port ListsSoap port, connected with SharePoint
     * @throws Exception in case of invalid paramaters or connection error.
     */
    public static ListsSoap sharePointListsAuth(String userName, String password, String cookieToken) throws Exception {

        ListsSoap port = null;


        if (userName != null && password != null && instanced) {
            try {
                Lists service = new Lists(new URL(_wsdlURL), new QName(_soapURL, "Lists"));
                port = service.getListsSoap();
                _out.log(Level.INFO, "Web Service Auth Username {0}", userName);
                ((BindingProvider) port).getRequestContext().put(BindingProvider.USERNAME_PROPERTY, userName);
                ((BindingProvider) port).getRequestContext().put(BindingProvider.PASSWORD_PROPERTY, password);
                ((BindingProvider) port).getRequestContext().put(BindingProvider.ENDPOINT_ADDRESS_PROPERTY, _wsdlURL.split("\\?")[0]);

                //Adding cookies support (obtained from claims-based authentication methods)
                //portions of the code belows to http://java.net/jira/browse/JAX_WS-1044
                Map<String, List<String>> reqHeaders = new HashMap<String, List<String>>();
                List<String> list = new ArrayList<String>();
                list.add(cookieToken);
                reqHeaders.put("Cookie", list);
                ((BindingProvider) port).getRequestContext().put(MessageContext.HTTP_REQUEST_HEADERS, reqHeaders);

            } catch (Exception e) {
                throw new Exception(e.getMessage());
            }
        } else {
            if (!instanced) {
                throw new Exception(Manager.class.getName() + ".createManagerService(String endPoint, String relPathToListLocation) must be the first execution.");
            } else {
                throw new Exception("Couldn't authenticate: Invalid connection details given.");
            }
        }

        return port;
    }

    /**
     * Creates a string from a XML file with start and end indicators
     *
     * @param docToString document to convert
     * @return string of the xml document
     */
    public static String xmlToString(Document docToString) {

        String returnString = "\n---------------- XML START ----------------\n";

        try {
            //create string from xml tree
            //Output the XML
            //set up a transformer
            TransformerFactory transfac = TransformerFactory.newInstance();
            Transformer trans;// = null;
            trans = transfac.newTransformer();
            trans.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
            trans.setOutputProperty(OutputKeys.INDENT, "yes");
            StringWriter sw = new StringWriter();
            StreamResult streamResult = new StreamResult(sw);
            DOMSource source = new DOMSource(docToString);
            trans.transform(source, streamResult);
            //print the XML
            returnString = returnString + sw.toString();
        } catch (TransformerException ex) {
            _out.log(Level.SEVERE, ex.getMessage());
        }

        returnString = returnString + "---------------- XML END ----------------\n";
        return returnString;
    }

    /**
     * Connect to a SharePoint List Web Service through the given open port, and
     * read all the elements of the given list. Only the ID and the given
     * attributes (column names) are displayed, as well as a dump of the SOAP
     * response from the Web Service (for debugging purposes).
     *
     * @param port an already authentificated SharePoint Online SOAP port
     * @param listName original name of the SharePoint list that is going to be
     * read
     * @param listColumnNames arraylist containing the various names of the
     * Columns of the Sharepoint list that are going to be read. If the column
     * name isn't found then an exception will be thrown
     * @param rowLimit limits the number of rows (list items) that are going to
     * be returned
     * @throws Exception
     */
    public static void displaySharePointList(ListsSoap port, String listName,
            ArrayList<String> listColumnNames, String rowLimit) throws Exception {

        if (port != null && listName != null && listColumnNames != null /*
                 * && rowLimit != null
                 */) {
            try {
                //Here are additional parameters that may be set
                String viewName = "";
                GetListItems.ViewFields viewFields = null;
                GetListItems.Query query = null;
                GetListItems.QueryOptions queryOptions = null;
                String webID = "";

                //Calling the List Web Service
                GetListItemsResponse.GetListItemsResult result = port.getListItems(listName,
                        viewName, query, viewFields, rowLimit, queryOptions, webID);
                Object listResult = result.getContent().get(0);
                if ((listResult != null) && (listResult instanceof ElementNSImpl)) {
                    ElementNSImpl node = (ElementNSImpl) listResult;

                    //Dumps the retrieved info in the console
                    Document document = node.getOwnerDocument();

                    _out.log(Level.OFF, "SharePoint Online Lists Web Service Response {0}", Manager.xmlToString(document));

                    //selects a list of nodes which have z:row elements
                    NodeList list = node.getElementsByTagName("z:row");
                    //list.getLength();

                    //Displaying every result received from SharePoint, with its ID
                    for (int i = 0; i < list.getLength(); i++) {

                        //Gets the attribute of the current row/element
                        NamedNodeMap attributes = list.item(i).getAttributes();
                        System.out.print("id: " + attributes.getNamedItem("ows_ID").getNodeValue() + ",\t");

                        for (String columnName : listColumnNames) {
                            String internalColumnName = "ows_" + columnName;
                            if (attributes.getNamedItem(internalColumnName) != null) {
                                System.out.println("column_name: " + columnName + ", value: " + attributes.getNamedItem(internalColumnName).getNodeValue());
                            }
                        }
                    }
                } else {
                    throw new Exception(listName + " list response from SharePoint is either null or corrupt\n");
                }
            } catch (Exception ex) {
                throw new Exception(ex.getMessage());
            }
        }
    }

    /**
     * This portion of code, returns a Node Object which is usefull for the
     * GetListItems.Query object. Code belongs to Leonid Sokolin
     * http://davidsit.wordpress.com/2010/02/10/reading-a-sharepoint-list-with-java-tutorial/#comment-66
     *
     * @param sXML
     * @return
     * @throws ParserConfigurationException
     * @throws SAXException
     * @throws IOException
     */
    public static Node generateXmlNode(String sXML) throws ParserConfigurationException,
            SAXException, IOException {

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setValidating(false);
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document documentOptions = builder.parse(new InputSource(new StringReader(sXML)));
        Node elementOptions = documentOptions.getDocumentElement();

        return elementOptions;
    }

    /**
     * This method is for using other methods
     *
     * @param port
     * @param listName
     * @param ows_ID
     * @return
     * @throws Exception
     */
    public static String getUIDFromListElement(ListsSoap port, String listName, String ows_ID) throws Exception {

        String ret = null;
        if (port != null && listName != null && ows_ID != null) {
            //Here are additional parameters that may be set
            String viewName = "";
            GetListItems.ViewFields viewFields = null;
            GetListItems.Query query = new GetListItems.Query();
            String qry = "<Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + ows_ID + "</Value></Eq></Where></Query>";
            query.getContent().add(generateXmlNode(qry));
            GetListItems.QueryOptions queryOptions = null;
            String webID = "";
            //Calling the List Web Service
            GetListItemsResponse.GetListItemsResult result = port.getListItems(listName,
                    viewName, query, viewFields, "1", queryOptions, webID);
            Object listResult = result.getContent().get(0);
            if ((listResult != null) && (listResult instanceof ElementNSImpl)) {
                ElementNSImpl node = (ElementNSImpl) listResult;

                //Dumps the retrieved info in the console
                //Document document = node.getOwnerDocument();
                //_out.log(Level.OFF, "SharePoint Online Lists Web Service Response {0}", Manager.xmlToString(document));

                //selects a list of nodes which have z:row elements
                NodeList list = node.getElementsByTagName("z:row");
                //list.getLength();

                //Displaying every result received from SharePoint, with its ID
                for (int i = 0; i < list.getLength(); i++) {

                    //Gets the attribute of the current row/element
                    NamedNodeMap attributes = list.item(i).getAttributes();
                    if (attributes.getNamedItem("ows_ID").getNodeValue().equals(ows_ID)) {
                        //ie. ows_UniqueId="3;#{18CD30B0-9384-4378-BFA8-5DA0EC87E2C0}"
                        ret = attributes.getNamedItem("ows_UniqueId").getNodeValue().split(";#")[1];

                    }
                }
            } else {
                throw new Exception(listName + " list response from SharePoint is either null or corrupt\n");
            }
        }

        return ret;
    }

    /**
     * CAML query can be found at:
     * http://stackoverflow.com/questions/1671664/accessing-last-record-in-a-sharepoint-list-via-caml-query
     * Obtain the GUID of a last list item
     *
     * @param port
     * @param listName
     * @return the ID, from the ows_UniqueId - ie.
     * "3;#{18CD30B0-9384-4378-BFA8-5DA0EC87E2C0}" - of the item
     * @throws Exception
     */
    public static String getIdFromLastListElement(ListsSoap port, String listName)
            throws ParserConfigurationException, SAXException, IOException, Exception {

        String ret = null;

        if (port != null && listName != null) {
            //Here are additional parameters that may be set
            String viewName = "";
            GetListItems.ViewFields viewFields = null;
            GetListItems.Query query = new GetListItems.Query();
            //String qry = "<View><RowLimit>1</RowLimit><Query><OrderBy><FieldRef Name='ID' Ascending='False' /></OrderBy></Query></View>";
            String qry = "<Query><OrderBy><FieldRef Name='ID' Ascending='False' /></OrderBy></Query>";
            query.getContent().add(generateXmlNode(qry));
            GetListItems.QueryOptions queryOptions = null;
            String webID = "";
            //Calling the List Web Service
            GetListItemsResponse.GetListItemsResult result = port.getListItems(listName,
                    viewName, query, viewFields, "1", queryOptions, webID);
            Object listResult = result.getContent().get(0);
            if ((listResult != null) && (listResult instanceof ElementNSImpl)) {
                ElementNSImpl node = (ElementNSImpl) listResult;

                //selects a list of nodes which have z:row elements
                NodeList list = node.getElementsByTagName("z:row");
                NamedNodeMap attributes = list.item(0).getAttributes();
                ret = attributes.getNamedItem("ows_UniqueId").getNodeValue().split(";#")[0];
            } else {
                throw new Exception(listName + " list response from SharePoint is either null or corrupt\n");
            }
        }

        return ret;
    }

    /**
     * Obtain the GUID of a List given its name
     *
     * @see {@link #addAttachMentToListItem(com.microsoft.schemas.sharepoint.soap.ListsSoap,
     * java.lang.String, java.lang.String, java.lang.String, java.lang.String)
     * @param port the SOAP port
     * @param listName the name of the list
     * @return the GUID of the list
     */
    public static String getUIDFromList(ListsSoap port, String listName) {

        String ret = null;

        GetListResult result = port.getList(listName);
        Object obj = result.getContent().get(0);

        if ((obj != null) && (obj instanceof ElementNSImpl)) {
            ElementNSImpl node = (ElementNSImpl) obj;
            ret = node.getAttribute("ID");
            //System.out.println("list_name: " + node.getAttribute("Name"));
            //System.out.println("list_id: " + ret);
        }
        return ret;
    }

    /**
     * This function will insert the given item in the SharePoint that
     * corresponds to the list name given (or list GUID).
     *
     * @param port an already authentificated SharePoint SOAP port
     * @param listName SharePoint list name or list GUID (guid must be enclosed
     * in braces)
     * @param itemAttributes This represents the content of the item that need
     * to be inserted. The key represents the type of attribute (SharePoint
     * column name) and the value corresponds to the item attribute value.
     *
     * this code belongs to: http://davidsit.wordpress.com/tag/connector/
     */
    public static void insertListItem(ListsSoap port, String listName, HashMap<String, String> itemAttributes) throws Exception {

        //Parameters validity check
        if (port != null && listName != null && itemAttributes != null && !itemAttributes.isEmpty()) {
            try {

                //Building the CAML query with one item to add, and printing request
                ListsRequest newCompanyRequest = new ListsRequest("New");
                newCompanyRequest.createListItem(itemAttributes);

                //initializing the Web Service operation here
                Updates updates = new UpdateListItems.Updates();

                //Preparing the request for the update
                Object docObj = (Object) newCompanyRequest.getRootDocument().getDocumentElement();
                updates.getContent().add(0, docObj);
                //Sending the insert request to the Lists.UpdateListItems Web Service
                /*
                 * UpdateListItemsResult result =
                 */
                UpdateListItemsResult updateListItems = port.updateListItems(listName, updates);
                com.sun.org.apache.xerces.internal.dom.ElementNSImpl get = (com.sun.org.apache.xerces.internal.dom.ElementNSImpl) updateListItems.getContent().get(0);
                //com.sun.org.apache.xerces.internal.dom.ElementNSImpl
                Manager.xmlToString(get.getOwnerDocument());

            } catch (Exception e) {
                throw new Exception(e.getMessage());
                // e.printStackTrace();
            }
        }
    }

    /**
     * This method attach an element to a specific list item - listID -
     *
     * @param port the SOAP port.
     * @param GUIDList the unique identifier of the list ie.
     * {18CD30B0-9384-4378-BFA8-5DA0EC87E2C0}
     * @see {@link #getUIDFromList(com.microsoft.schemas.sharepoint.soap.ListsSoap, java.lang.String)
     * }.
     * @param listItemID the number of the list item.
     * @param filePath the path of the file that is going to be stored.
     * @param fileName the name of the file that is going to be stored.
     * @return a String containing the relative path of the file.
     *
     * @see
     * http://stackoverflow.com/questions/1264709/convert-inputstream-to-byte-in-java
     * @see http://commons.apache.org/io/
     */
    public static String addAttachMentToListItem(ListsSoap port, String GUIDList, String listItemID, String filePath, String fileName) throws IOException {

        String addAttachment = null;
        if (port != null && GUIDList != null && listItemID != null && filePath != null && fileName != null) {
            addAttachment = port.addAttachment(GUIDList, listItemID, fileName,
                    IOUtils.toByteArray(new FileInputStream(filePath)));
        }

        return addAttachment;
    }

    /**
     * This method attach an element to a specific list item - listID -
     *
     * @param port the SOAP port.
     * @param GUIDList the unique identifier of the list ie.
     * {18CD30B0-9384-4378-BFA8-5DA0EC87E2C0}
     * @see {@link #getUIDFromList(com.microsoft.schemas.sharepoint.soap.ListsSoap, java.lang.String)
     * }.
     * @param listItemID the number of the list item.
     * @param filePath the path of the file that is going to be stored.
     * @param inputStream the stream of the file to be stored.
     * @return a String containing the relative path of the file.
     *
     * @see
     * http://stackoverflow.com/questions/1264709/convert-inputstream-to-byte-in-java
     * @see http://commons.apache.org/io/
     */
    public static String addAttachMentToListItem(ListsSoap port, String GUIDList, String listItemID, String fileName, InputStream inputStream) throws IOException {

        String addAttachment = null;
        if (port != null && GUIDList != null && listItemID != null && fileName != null && inputStream != null) {
            addAttachment = port.addAttachment(GUIDList, listItemID, fileName,
                    IOUtils.toByteArray(inputStream));
        }

        return addAttachment;
    }

    /**
     * This method creates a list item and attaches an file to it given a list
     * name
     *
     * @param port
     * @param listName
     * @param item
     * @param filePath
     * @param fileName
     * @throws Exception
     */
    public static void addItemWithAttachmentToList(ListsSoap port, String listName, HashMap<String, String> item, String filePath, String fileName) throws Exception {

        String _guidList = Manager.getUIDFromList(port, listName);
        System.out.println("List ID Returned: " + _guidList);

        Manager.insertListItem(port, listName, item);
        String ows_ID = Manager.getIdFromLastListElement(port, listName);
        System.out.println("Item ID Returned: " + ows_ID);

        String addAttachment = Manager.addAttachMentToListItem(port, _guidList, ows_ID, filePath, fileName);
        System.out.println("Attachment PATH Returned: " + addAttachment);
    }

    /**
     * This method creates a list item and attaches an file to it given a list
     * name
     *
     * @param port
     * @param listName
     * @param item
     * @param fileName
     * @param inputStream
     * @throws Exception
     */
    public static void addItemWithAttachmentToList(ListsSoap port, String listName, HashMap<String, String> item, String fileName, InputStream inputStream) throws Exception {

        String _guidList = Manager.getUIDFromList(port, listName);
        System.out.println("List ID Returned: " + _guidList);

        Manager.insertListItem(port, listName, item);
        String ows_ID = Manager.getIdFromLastListElement(port, listName);
        System.out.println("Item ID Returned: " + ows_ID);

        String addAttachment = Manager.addAttachMentToListItem(port, _guidList, ows_ID, fileName, inputStream);
        System.out.println("Attachment PATH Returned: " + addAttachment);
    }
}