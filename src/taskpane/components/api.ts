export default { getFields, insertField };

export interface Field {
  name: string;
  uniqueName: string;
  description: string;
  displayName: string;
  condition: string;
  programmeType: string;
  outputText: string;
  example?: string;
  dataType?: string;
}

async function fetchFile(url: string) {
  try {
    const resp = await fetch(url);
    return await resp.text();
  } catch (err) {
    console.log("error fetching", url, err);
    throw err;
  }
}

async function getFields() {
  // var schemaId = Office.context.document.settings.get("schema");
  console.log("init started");
  const [schema, extendedSchema] = await Promise.all([
    fetchFile("/assets/schema.xml"),
    fetchFile("/assets/extendedschema.xml"),
  ]);
  return await new Promise<Field[]>((resolve, reject) => {
    // Office.context.document.customXmlParts.addAsync(schema, (result) => {
    //   console.log("adding schema", result.value);
    //   //save the custom schema id in settings so that we can retrive it later
    //   Office.context.document.settings.set("schema", result.value.id);
    //   //this will persist the above settings inside document
    //   Office.context.document.settings.saveAsync((asyncResult) => {
    //     console.log(`Settings saved with status: ${asyncResult.status}`);
    //   });
    // });
    Office.context.document.customXmlParts.addAsync(extendedSchema, () => {
      Office.context.document.customXmlParts.getByNamespaceAsync(
        "http://kleash.github.io/extendeddata",
        (schemaRes) => {
          if (schemaRes.error) {
            reject(schemaRes.error);
            return;
          }
          const xmlSchema = schemaRes.value[0];
          const domParser = new DOMParser();
          xmlSchema.getNodesAsync("*", (nodesRes) => {
            if (nodesRes.error) {
              reject(nodesRes.error);
              return;
            }
            const promises: Promise<Field[]>[] = [];
            nodesRes.value.forEach((node) => {
              console.log("processing root node", node);
              promises.push(
                new Promise((resolve, reject) => {
                  const fields = [];
                  node.getXmlAsync((nodeRes) => {
                    if (nodeRes.error) {
                      reject(nodeRes.error);
                      return;
                    }
                    const xml = nodeRes.value;
                    const dom = domParser.parseFromString(xml, "text/xml").getElementsByTagName("extendeddata")[0];
                    if (!dom) {
                      reject('something is wrong, element "extendeddata" not existed: ' + xml);
                      return;
                    }
                    try {
                      dom.childNodes.forEach((child) => {
                        const field = nodeToField(dom, child);
                        if (field) fields.push(field);
                      });
                    } catch (e) {
                      reject(e);
                      return;
                    }
                    resolve(fields);
                  });
                })
              );
            });
            console.log("promises", promises);
            Promise.all(promises)
              .then((results) => results.reduce((p, e) => p.concat(e), []))
              .then((fields) => resolve(fields))
              .catch((err) => reject(err));
          });
        }
      );
    });
    Office.context.document.customXmlParts.addAsync(
      '<?xml version="1.0"?><conditions xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://opendope.org/conditions"/>'
    );
    Office.context.document.customXmlParts.addAsync(
      '<?xml version="1.0"?><components xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://opendope.org/components"/>'
    );
    Office.context.document.customXmlParts.addAsync(
      '<?xml version="1.0"?><xpaths xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://opendope.org/xpaths"/>'
    );
  });
}

function nodeToField(parent: Element, child: ChildNode): Field {
  console.debug("nodeToField", child);
  var displayName = "";
  var description = "";
  var tagName = "";
  var programmeType = "";
  var condition = "";
  var childName = child.nodeName;
  var xpathChild = parent.getElementsByTagName(childName)[0];
  if (!xpathChild) {
    return null;
  }
  var displayNameTag = xpathChild.getElementsByTagName("displayName");
  if (displayNameTag.length > 0) {
    displayName = displayNameTag[0].childNodes[0].nodeValue;
  }
  var descriptionTag = xpathChild.getElementsByTagName("description");
  if (descriptionTag != null && descriptionTag.length > 0) {
    description = descriptionTag[0].childNodes[0].nodeValue;
  }
  var tagNameTag = xpathChild.getElementsByTagName("tagName");
  if (tagNameTag != null && tagNameTag.length > 0) {
    tagName = tagNameTag[0].childNodes[0].nodeValue;
  }
  var programmeTypeTag = xpathChild.getElementsByTagName("programmeType");
  if (programmeTypeTag != null && programmeTypeTag.length > 0) {
    programmeType = programmeTypeTag[0].childNodes[0].nodeValue;
  }
  var conditionTag = xpathChild.getElementsByTagName("condition");
  if (conditionTag != null && conditionTag.length > 0) {
    condition = conditionTag[0].childNodes[0].nodeValue;
  }
  return {
    name: tagName,
    outputText: `{ ${tagName} }`,
    description,
    condition,
    displayName,
    programmeType,
    uniqueName: "/fixmarketplace[1]/" + childName + "[1]",
  };
}

//Add annotated field in document and customXml with namespace as Xpaths
function insertField(field: Field) {
  console.log("insertField", field);
  if (!field) {
    return;
  }
  //TODO selected field
  var elementToAdd = field.uniqueName;
  var phText = field.outputText;

  //Generate od
  var uid = makeid(5);
  var schemaId = Office.context.document.settings.get("schema");

  // STEP 2: add to xpaths xml
  //get xml schema from namespace
  Office.context.document.customXmlParts.getByNamespaceAsync("http://opendope.org/xpaths", (result) => {
    var xmlPart = result.value[0];
    //get root node of xml schema
    xmlPart.getNodesAsync("*", (nodeResults) => {
      for (var i = 0; i < nodeResults.value.length; i++) {
        var node = nodeResults.value[i];
        //get xml of root node and add the xpath
        node.getXmlAsync((result) => {
          //parse the xml
          var parser = new DOMParser();
          var xml = parser.parseFromString(result.value, "text/xml");
          var parent = xml.getElementsByTagName("xpaths")[0];

          //TODO can we execute at startup and keep uid:xpath in some map? This will remove the extra step

          // EXTRA STEP : check if already exist in xpath, if yes, use that uid and skip "add to xpath" step
          //all xpath
          var containsXpath = false;
          for (var i = 0; i < parent.getElementsByTagName("xpath").length; ++i) {
            var xpathChild = parent.getElementsByTagName("xpath")[i];
            var databindingChild = xpathChild.getElementsByTagName("dataBinding");
            if (databindingChild != null && databindingChild.length > 0) {
              if (databindingChild[0].getAttribute("xpath") == elementToAdd) {
                uid = xpathChild.getAttribute("id");
                containsXpath = true;
              }
            }
          }

          //Add new xpath to document
          if (!containsXpath) {
            //add a xpath child, this id is used in tag of content control
            var newEle = xml.createElementNS(parent.namespaceURI, "xpath");

            newEle.setAttribute("id", uid);

            //add sub child of xpath i.e. dataBinding which have xpath value and pointer to datasource i.e. our schema xml
            var dataBindingEle = xml.createElementNS(parent.namespaceURI, "dataBinding");
            dataBindingEle.setAttribute("xpath", elementToAdd);
            dataBindingEle.setAttribute("storeItemID", schemaId);
            newEle.appendChild(dataBindingEle);
            parent.appendChild(newEle);

            //convert xml back to string
            var s = new XMLSerializer();
            var newXmlStr = s.serializeToString(xml);

            //======= Update Schema back to xml ===============
            //TODO Replace with updateSchema(newXmlStr) function
            //get schema id from settings
            schemaId = Office.context.document.settings.get("schema");
            //get xml schema from id
            Office.context.document.customXmlParts.getByNamespaceAsync("http://opendope.org/xpaths", (result) => {
              var xmlPart = result.value[0];
              //get root node of xml schema
              xmlPart.getNodesAsync("*", (nodeResults) => {
                for (var i = 0; i < nodeResults.value.length; i++) {
                  var node = nodeResults.value[i];
                  //set back the schema
                  node.setXmlAsync(newXmlStr);
                }
              });
            });
          }
          //======= Update Schema back to xml END===============

          // STEP 3: create normal content control
          Word.run((context) => {
            // Queue commands to create a content control.
            var serviceNameRange = context.document.getSelection();
            var serviceNameContentControl = serviceNameRange.insertContentControl();
            //get values from text box to set as property of content control

            //Removing title as placeholder is enough and it will keep it cleaner
            serviceNameContentControl.title = "";
            serviceNameContentControl.tag = "od:xpath=" + uid;
            serviceNameContentControl.placeholderText = phText;
            serviceNameContentControl.appearance = "Hidden";
            serviceNameContentControl.color = "blue";

            return context.sync();
          }).catch((error) => {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
              console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
          });
        });
      }
    });
  });
}

function makeid(length) {
  var result = "";
  var characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  var charactersLength = characters.length;
  for (var i = 0; i < length; i++) {
    result += characters.charAt(Math.floor(Math.random() * charactersLength));
  }
  return result;
}
