export default { getFields, insertField, registerSelectionListener };

export interface Field {
  name: string;
  uniqueName: string;
  displayName: string;
  description: string;
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

async function getFields(overwite?: boolean) {
  return loadSchema(overwite).then(() => loadFields());
}

async function loadSchema(overwriteSchema: boolean): Promise<boolean> {
  const schemaId = Office.context.document.settings.get("schema");
  if (schemaId) {
    console.log("schema existed", schemaId);
    if (!overwriteSchema) {
      return true;
    }
    console.log("delete existing schema", schemaId);
    await deleteSchema(schemaId);
  }
  const [schema, extendedSchema] = await Promise.all([
    fetchFile("/assets/schema.xml"),
    fetchFile("/assets/extendedschema.xml"),
  ]);
  // create schema
  const promises: Promise<boolean>[] = [];
  promises.push(
    new Promise((resolve) => {
      Office.context.document.customXmlParts.addAsync(schema, ({ value }) => {
        Office.context.document.settings.set("schema", value.id);
        Office.context.document.settings.saveAsync();
        resolve(!!value);
      });
    })
  );
  promises.push(
    new Promise((resolve) => {
      Office.context.document.customXmlParts.addAsync(extendedSchema, ({ value }) => resolve(!!value));
    })
  );
  promises.push(
    new Promise((resolve) => {
      Office.context.document.customXmlParts.addAsync(
        '<?xml version="1.0"?><components xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://opendope.org/components"/>',
        ({ value }) => resolve(!!value)
      );
    })
  );
  promises.push(
    new Promise((resolve) => {
      Office.context.document.customXmlParts.addAsync(
        '<?xml version="1.0"?><xpaths xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://opendope.org/xpaths"/>',
        ({ value }) => resolve(!!value)
      );
    })
  );
  return Promise.all(promises).then((values) => values.every((e) => e));
}

async function loadFields(): Promise<Field[]> {
  return new Promise((resolve, reject) => {
    Office.context.document.customXmlParts.getByNamespaceAsync("http://kleash.github.io/extendeddata", (schemaRes) => {
      if (schemaRes.error) {
        reject(schemaRes.error);
        return;
      }
      try {
        const xmlSchema = schemaRes.value[0];
        xmlSchema.getNodesAsync("*", (nodesRes) => {
          if (nodesRes.error) {
            reject(nodesRes.error);
            return;
          }
          nodesRes.value[0].getXmlAsync((nodeRes) => {
            if (nodeRes.error) {
              reject(nodeRes.error);
              return;
            }
            const xml = nodeRes.value;
            const domParser = new DOMParser();
            const dom = domParser.parseFromString(xml, "text/xml").getElementsByTagName("extendeddata")[0];
            if (!dom) {
              reject('something is wrong, element "extendeddata" not found');
              return;
            }
            const map = new Map<string, Field>();
            dom.childNodes.forEach((child) => {
              const field = nodeToField(dom, child);
              if (field) {
                if (map.has(field.name)) {
                  console.warn("field already existed", map.get(field.name), field);
                } else {
                  map.set(field.name, field);
                }
              }
            });
            resolve(Array.from(map.values()));
          });
        });
      } catch (e) {
        reject(e);
      }
    });
  });
}

async function deleteSchema(schemaId: string) {
  const promises: Promise<any>[] = [];
  promises.push(
    new Promise((resolve) => {
      Office.context.document.customXmlParts.getByIdAsync(schemaId, (rs: Office.AsyncResult<Office.CustomXmlPart>) => {
        rs.value?.deleteAsync(() => {
          console.log("deleted schema");
          resolve(1);
        });
      });
    })
  );
  promises.push(
    new Promise((resolve) => {
      Office.context.document.customXmlParts.getByNamespaceAsync(
        "http://opendope.org/xpaths",
        (rs: Office.AsyncResult<Office.CustomXmlPart[]>) => {
          if (rs.value && rs.value[0]) {
            rs.value[0].deleteAsync(() => {
              console.log("deleted xpaths");
              resolve(1);
            });
          }
        }
      );
    })
  );
  promises.push(
    new Promise((resolve) => {
      Office.context.document.customXmlParts.getByNamespaceAsync(
        "http://opendope.org/components",
        (rs: Office.AsyncResult<Office.CustomXmlPart[]>) => {
          if (rs.value && rs.value[0]) {
            rs.value[0].deleteAsync(() => {
              console.log("deleted components");
              resolve(1);
            });
          }
        }
      );
    })
  );
  promises.push(
    new Promise((resolve) => {
      Office.context.document.customXmlParts.getByNamespaceAsync(
        "http://kleash.github.io/extendeddata",
        (rs: Office.AsyncResult<Office.CustomXmlPart[]>) => {
          if (rs.value && rs.value[0]) {
            rs.value[0].deleteAsync(() => {
              console.log("deleted extendeddata");
              resolve(1);
            });
          }
        }
      );
    })
  );
  return Promise.all(promises);
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
  var uid = ((length) => {
    var result = "";
    var characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    var charactersLength = characters.length;
    for (var i = 0; i < length; i++) {
      result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }
    return result;
  })(5);

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
            const control = context.document.getSelection().insertContentControl();
            control.tag = "od:xpath=" + uid;
            control.placeholderText = phText;
            control.appearance = "Hidden";

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

function registerSelectionListener(handler: (fieldId: string) => void) {
  const selectionHandler = (event: Office.DocumentSelectionChangedEventArgs) => {
    event.document.getSelectedDataAsync(Office.CoercionType.Text, null, (selectionRs) => {
      console.debug("doc selection", selectionRs.value);
      const text = selectionRs.value as string;
      // TODO: can this be replaced by a more robust logic, e.g. get the field id from its xml data?
      // Regex: "{ any value but not curly brackets }"
      if (/{\s([^{}])+\s}/.test(text)) {
        handler(text.substring(2, text.length - 2));
      } else {
        handler(null);
      }
    });
  };
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectionHandler);
}

function nodeToField(parent: Element, child: ChildNode): Field {
  var displayName = "";
  var description = "";
  var name = "";
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
    name = tagNameTag[0].childNodes[0].nodeValue;
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
    name,
    displayName,
    description,
    condition,
    programmeType,
    outputText: `{ ${displayName} }`,
    uniqueName: "/fixmarketplace[1]/" + childName + "[1]",
  };
}
