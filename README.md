# Custom XML Parts

_This branch is dedicated to describing the custom XML part design_


Microsoft Office documents allow users to embed custom XML data as part of the document. This data is named a custom XML part.

You can create and modify custom XML parts in a document by using the APIs described below.

Usage:

Most of the XML parts in the document are built-in parts that help to define the structure and the state of the document. However, documents can also contain custom XML parts, which you can use to store arbitrary XML data in the documents.

The XML file formats enable applications to extend the document by stroing add-in usage, state, or configuration data as part of the document. Since it travels with the document, it is very convenient to understand the state of the document.

Current JavaScript API support:

Currently, only Word supports custom XML part through the JavaScript 1.0 API set. It allows XML part and node level access with XML part level event support.

Addition:

We are adding custom XML part API to the modern 1.1+ API set for Word, Excel and PowerPoint. Excel and Word will get the new APIs first followed by PowerPoint. This open spec is about the new custom XML part API design.

## Updated object model
(this sample is provided for Excel. Similar pattern would apply for Word as well.)

```typescript
// Addition of custom XML parts collection to workbook object.
/**
 * Represents the collection of custom XML parts contained by this workbook. Read-only.
 */
customXmlParts: CustomXmlPartCollection

// Declare custom xml parts collection

/**
 * A collection of custom XML parts.
 */
declare class CustomXmlPartCollection extends IEnumerable < CustomXmlPart > {

    /**
     * Gets a custom XML part based on its ID.
     * @param id Id of the custom XML part to be returned
     * @returns CustomXmlPart
     */
    getItem(id: string): CustomXmlPart;

    /**
     * Gets a custom XML part based on its ID. Returns null-object if the ID is not present.
     * @param id Id of the custom XML part to be returned
     * @returns CustomXmlPart
     */
    getItemOrNullObject(id: string): CustomXmlPart;

    /**
     * Gets the number of items in the collection.
     */
    getCount(): integer;

    /**
     * Adds a new custom XML part to the workbook.
     * @param xml XML content as string. Must be a valid XML fragment.
     * @returns CustomXmlPart
     */
    add(xmlContent: string): CustomXmlPart;

    /**
     * Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.
     * @param namespaceUri Namespace string to be searched
     * @returns CustomXmlPartScopedCollection
     */
    getByNamespace(namespaceUri: string): CustomXmlPartScopedCollection;
}


/**
 * A scoped collection of custom XML parts. A scoped collection is the result of some operation, e.g. filtering by namespace. A scoped collection cannot be scoped any further.
 */

declare class CustomXmlPartScopedCollection extends IEnumerable < CustomXmlPart > {
    /**
     * Gets a custom XML part based on its ID.
     * @param id Id of the custom XML part to be returned
     * @returns CustomXmlPart
     */
    getItem(id: string): CustomXmlPart;

    /**
     * Gets a custom XML part based on its ID only if there is single part present.  
     * Usage: parts.getByNamespace('ns').getItemOnly('id') // this will return the part if there is only 1 part on the getByNamespace('ns') response. Else it throws.
     * @param id Id of the custom XML part to be returned
     * @returns CustomXmlPart
     */
    getItemOnly(id: string): CustomXmlPart;

    /**
     * Gets a custom XML part based on its ID. Returns null-object if the ID is not present.
     * Usage: parts.getByNamespace('ns').getItemOnly('id') // this will return the part if there is only 1 part on the getByNamespace('ns') response. Else it returns null object.     
     * @param id Id of the custom XML part to be returned
     * @returns CustomXmlPart
     */
    getOnlyItemOrNullObject(id: string): CustomXmlPart;

    /**
     * Gets the number of items in the collection.
     */
    getCount(): integer;

    /**
     * The custom XML part collection's namespace URI. Read-only.
     */
    namespaceUri: string
}

// CustomXmlPart OBJECT

/**
 * Represents a custom XML part object in a workbook.
 */
declare class CustomXmlPart {
    /**
     * Deletes the custom XML part.
     * @returns void
     */
    delete(): void;

    /**
     * The custom XML part's ID. Read-only.
     */
    id: string;

    /**
     * The custom XML part's namespace URI. Read-only.
     */
    namespaceUri: string

    /**
     * Returns the XML content as a string.
     * @returns string
     */
    getXml(): string

    /**
     * Sets the entire XML part
     * @returns void
     */
    setXml(xml: string): void

    /**
     * Inserts an XML node under the parent element identified using xpath query.
     * @param xpath The Xpath used to determine the parent element under which the new element will be inserted. This must resolve to a single node.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @param xml Actual XML element to be inserted
     * @param index The position where the new element should be inserted under the parent node. By default the new element will be added at the end. 0-index based.
     * @returns void
     */
    insertElement(xpath: string, xml: string, namespaceMappings?: Object, index ? : int): void

    /**
     * Update the XML element using the new xml string.
     * @param xpath The Xpath used to determine element to be updated. This must resolve to a single node.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @param xml Actual XML element that is used to update the target node.
     * @returns void
     */
    updateElement(xpath: string, xml: string, namespaceMappings?: Object): void

    /**
     * Delete a XML element.
     * @param xpath The Xpath used to determine element to be deleted. This must resolve to a single node.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @returns void
     */
    deleteElement(xpath: string, namespaceMappings?: Object): void

    /**
     * Query the XML document and return 1 or more XML elements.
     * @param xpath The Xpath used to search for the elements.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @returns string[] 
     */
    query(xpath: string, namespaceMappings?: Object): string[]

    /**
     * Insert XML attribute on the identified element.
     * @param xpath The Xpath used to locate for the target element. This must resolve to the target attribute.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @param name Name of the attribute to be added.
     * @param value Value of the attribute to be added.
     * @returns void
     */
    insertAttribute(xpath: string, namespaceMappings?: Object, name: string, value: string): void

    /**
     * Update XML attribute on the identified element.
     * @param xpath The Xpath used to locate for the target element. This must resolve to the target attribute.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @param value Value of the attribute to be updated.
     * @returns void
     */
    updateAttribute(xpath: string, value: string, namespaceMappings?: Object): void

    /**
     * Delete a XML attribute.
     * @param xpath The Xpath used to determine attribute to be deleted. This must resolve to the target attribute.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @returns void
     */
    deleteAttribute(xpath: string, namespaceMappings?: Object): void

    //
    //
    // XML NODE SECTION: Below APIs are for future expansion.
    //
    //


    /**
     * Get XML nodes associated with the part.
     * @param xpath The Xpath used to query.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @returns CustomXmlNodeCollection Collection of XML nodes
     */
    getXmlNodes(xpath: string, namespaceMappings?: Object): CustomXmlNodeCollection

    /**
     * Get a single XML node.
     * @param xpath The Xpath used to query that results in a single node.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @returns CustomXmlNode Single XML node
     */
    getXmlNode(xpath: string, namespaceMappings?: Object): CustomXmlNode

    /**
     * Get a single XML node.
     * @param xpath The Xpath used to query that results in a single node.
     * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
     * @returns CustomXmlNode Single XML node
     */
    getXmlNode(xpath: string, namespaceMappings?: Object): CustomXmlNode

}

//
//
//
// XML NODE SECTION: Below parts are for future expansion.
//
//
//

/**
 * Represents a collection of custom XML nodes
 */
declare class CustomXmlNodeCollection {
  /**
   * Gets a custom XML node based on its index position.
   * @param index
   * @returns CustomXmlPart
   */
  getItemAt(index: integer): CustomXmlNode;

  // Given that nodes don't have keys - we can't have getItem() for this object

  /**
   * Gets the number of items in the collection.
   */
  getCount(): integer;

  /**
   * Adds a new custom XML node
   * @param nodeContent
   * @returns CustomXmlNode
   */
  add(nodeContent: string): CustomXmlNode;

}

/**
 * Represents a node of XML document
 */
declare class CustomXmlNode {
    /**
     * An object that consists of {"name": "value"} mappings for each of the attribute belonging to a single node.
     */
    attributes: XmlAttributeCollection;

    /**
     * Insert XML node before the current node.
     * @param nodeContent XML node content to be inserted
     * @returns void
     */
    insertBefore(nodeContent: string): void

    /**
     * Insert XML node after the current node.
     * @param nodeContent XML node content to be inserted
     * @returns void
     */
    insertAfter(nodeContent: string): void

    /**
     * Update current node
     * @param nodeContent XML node content to be updated with  
     * @returns void
     */
    setcontent(nodeContent: string): void

    /**
     * Deletes the node.
     * @returns void
     */
    delete(): void;  
}

/**
 * Represents a collection of custom XML attributes
 */
declare class CustomXmlAttributeCollection {

  /**
   * Gets a custom XML attribute based on its key.
   * @param id key value
   * @returns CustomXmlAttribute
   */
  getItem(id: string): CustomXmlAttribute;

  /**
   * Gets a custom XML part attribute on its key. Returns null-object if the ID is not present.
   * @param id key value
   * @returns CustomXmlAttribute
   */
  getItemOrNullObject(id: string): CustomXmlAttribute;


  /**
   * Gets the number of items in the collection.
   */
  getCount(): integer;

  /**
   * Adds a new custom XML attribute  
   * @param attribute key-value pair object
   * @param namespaceMappings an object that consists of {"prefix": "namespace"} mappings used in the Xpath.
   * @returns CustomXmlAttribute
   */
  add(attribute: CustomXmlAttribute, namespaceMappings?: Object): CustomXmlAttribute;
}

/**
 * Represents an attribute on the XMl element
 */
declare class CustomXmlAttribute {

    /**
     * Key of the attribute
     */
    key: string;

    /**
     * Value of the attribute
     */
    key: value;

    /**
     * Deletes the attribute.
     * @returns void
     */
    delete(): void;  
}


```

## Examples

```typescript

// Get a collection custom XML parts and count

Excel.run(function (ctx) {
  var xmlparts = ctx.workbook.customXmlParts;
  xmlparts.load();
  return ctx.sync().then(function() {
    console.log('Count= ' + xmlparts.count)
    for (var i = 0; i < xmlparts.items.length; i++)
    {
      console.log(xmlparts.items[i].id);
      console.log(xmlparts.items[i].namespaceUri);
      console.log(xmlparts.items[i].xml);  //XML string

    }
  });
}).catch(function(error) { //...
});

// Get a part based in its id.
var xmlparts = ctx.workbook.customXmlParts.getItem('id');

// Get a collection of scoped custom XML part based in its namespace
var xmlscopedparts = ctx.workbook.customXmlParts.getByNamespace('https://foo.com/ns');

// Add a new XML part
ctx.workbook.customXmlParts.add(xml); // where XML is the actual XML part string.

// Delete a XML part
ctx.workbook.customXmlParts..getItem('id').delete()

// Delete a XML part
ctx.workbook.customXmlParts..getItem('id').delete()

/**
Note:
Developer constructs the XPATH by using any arbitary prefix and provides that new prefix:namespace mapping in the parameter. This will result in behind the scene API call before the actual API could be called: mapPrefixToNamespace('prefix', 'namespace') for each of the mapping used in the namespaceMapping object. This is a temporary mapping that lasts only for the session (or call?)

The namespace mappings are for the xpath only, they have no bearing at all on the xml string. Conversely, the xml string must be a valid stand-alone snippet of xml (meaning it must define all its own prefixes), and cannot rely on previously-defined namespaces from the destination XML nor can it use the ones defined in the namespace mappings supplied.
o	For example, this snippet from the .md file won’t work:
// Insert element
insertElement('(//ar:library/br:book)[1]', '<pr:price>$100</pr:price>',
{'ar': 'https://goo.com/library', 'br' : 'https://foo.com/book', 'pr' : 'https://foo.com/price'})
o	That’s because “<pr:price>$100</pr:price>” is not a valid snippet of XML. That’s because the namespace prefix ‘pr’ is not defined. The fact that ‘pr’ is defined in the namespace mappings does not affect this. The string of xml supplied as the parameter to insertElement (or any of the others) must be a valid stand-alone snippet, regardless of all other params passed. The namespace mappings are only consulted when resolving the xpath, and not when interpreting the in-coming xml snippet.

*/

// Insert element
insertElement('(//ar:library/br:book)[1]', '<pr:price xmlns:pr="https://foo.com/price">$100</pr:price>',
{'ar': 'https://goo.com/library', 'br' : 'https://foo.com/book'})

// Update element
updateElement('(//ar:library/br:book)[1]', '<pr:price xmlns:pr="https://foo.com/price">$100</pr:price>',
{'ar': 'https://goo.com/library', 'br' : 'https://foo.com/book'})

// Delete element
deleteElement('(//ar:library/br:book)[1]', {'ar': 'https://goo.com/library', 'br' : 'https://foo.com/book'})

// Query element
query('//ar:library/br:book', {'ar': 'https://goo.com/library', 'br' : 'https://foo.com/book'});

```

Sample custom XML document for reference:

```xml
<lib:library
   xmlns:lib="http://foo.com/ns/library"
   xmlns:hr="http://foo.com/ns/person">
  <lib:book id="b0836217462" available="true">
   <lib:isbn>0836217462</lib:isbn>
   <lib:title xml:lang="en">Being a Pet</lib:title>
   <hr:author id="CMS">
    <hr:name>Charles M Schulz</hr:name>
    <hr:born>1922-11-26</hr:born>
    <hr:dead>2000-02-12</hr:dead>
   </hr:author>
   <lib:character id="Snoopy">
    <hr:name>Snoopy</hr:name>
    <hr:born>1950-10-04</hr:born>
    <lib:qualification>extroverted beagle</lib:qualification>
   </lib:character>
   <lib:character id="Bloop">
    <hr:name>Lucy</hr:name>
    <hr:born>1952-03-03</hr:born>
    <lib:qualification>bossy, crabby and selfish</lib:qualification>
   </lib:character>
  </lib:book>
</lib:library>
```
**[Tell us what you think](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-CXP)**
