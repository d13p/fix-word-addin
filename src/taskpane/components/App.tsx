import { IComboBoxOption, Icon, Label, MessageBar, MessageBarType, ResponsiveMode, Text } from "@fluentui/react";
import { DefaultButton, IconButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import * as React from "react";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import officeApi, { Field } from "./api";
import styles from "./App.styles";
import { FieldPicker } from "./FieldPicker";

const products: IDropdownOption[] = [
  { key: "CD", text: "CD" },
  { key: "CP", text: "CP" },
  { key: "MTN", text: "MTN" },
];

interface FieldSelection extends IComboBoxOption, Field {}

export function App() {
  const [refresh, setRefresh] = React.useState(0);
  const [isLoading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string>(null);
  const [fields, setFields] = React.useState<FieldSelection[]>();
  const [selectedField, setSelectedField] = React.useState<FieldSelection>();
  const [selectedProducts, setSelectedProducts] = React.useState<string[]>([]);
  const filteredFields = React.useMemo(() => {
    console.debug("filteredFields", fields, products);
    return (fields || []).filter((e) => {
      return (
        selectedProducts == null ||
        !selectedProducts.length ||
        selectedProducts.some((p) => e.programmeType.includes(p))
      );
    });
  }, [fields, selectedProducts]);

  React.useEffect(() => {
    setLoading(true);
    setError(null);
    officeApi
      .getFields()
      .then((fields) => {
        console.log("loaded fields", fields);
        setFields(
          fields
            .sort((a, b) => a.name.localeCompare(b.name))
            .map((e) => ({
              key: e.name,
              text: e.name,
              ...e,
            }))
        );
      })
      .catch((err) => {
        setError(`Error loading schema. Error: ${err.message || err}`);
        console.error(err);
      })
      .finally(() => setLoading(false));
  }, [refresh]);

  const [selectedFieldId, setSelectedFieldId] = React.useState<string>(null);
  React.useEffect(() => {
    if (!fields || !selectedFieldId) {
      return;
    }
    const field = fields.find((e) => e.name === selectedFieldId);
    if (field) {
      setSelectedField(field);
    }
  }, [selectedFieldId, fields]);

  React.useEffect(() => {
    officeApi.registerSelectionListener((fieldId) => {
      setSelectedFieldId(fieldId);
    });
    return officeApi.removeSelectionListener;
  }, []);

  const menuProps: IContextualMenuProps = React.useMemo(
    () => ({
      items: [
        {
          key: "refresh",
          text: "Refresh tags",
          onClick: () => setRefresh((val) => ++val),
        },
        {
          key: "newItem",
          text: "Suggest new tags",
          disabled: true,
        },
        {
          key: "help",
          text: "Get help",
          disabled: true,
        },
      ],
    }),
    []
  );

  const onClear = React.useCallback(() => {
    setSelectedField(null);
    setSelectedProducts([]);
  }, []);

  const onInsert = React.useCallback(() => {
    officeApi.insertField(selectedField);
  }, [selectedField]);

  const onSelectField = React.useCallback((field) => {
    setSelectedField(field);
  }, []);

  const onSelectProduct = React.useCallback((_, item) => {
    if (item) {
      setSelectedProducts((products) => {
        return item.selected ? [...products, item.key as string] : products.filter((key) => key !== item.key);
      });
    }
  }, []);

  if (isLoading) {
    return (
      <div className={styles.center}>
        <Spinner label="Loading data..." size={SpinnerSize.large} />
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.center}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {error}
        </MessageBar>
      </div>
    );
  }

  return (
    <>
      <div className={styles.root}>
        <div className={styles.header}>
          <img src="assets/logo_inversed.svg" />
        </div>
        <div className={styles.body}>
          <div className={styles.product}>
            <Dropdown
              placeholder="Select products"
              onRenderTitle={(items) => {
                return (
                  <>
                    {items.map((item) => (
                      <span className={styles.selectedProduct}>{item.text}</span>
                    ))}
                  </>
                );
              }}
              options={products}
              selectedKeys={selectedProducts}
              onChange={onSelectProduct}
              multiSelect
              responsiveMode={ResponsiveMode.unknown}
            />
            <IconButton iconProps={{ iconName: "MoreVertical" }} menuProps={menuProps} onRenderMenuIcon={() => null} />
          </div>
          <div className={styles.tag}>
            <Label required>Select tag to annotate</Label>
            <FieldPicker fields={filteredFields} onSelect={onSelectField} />
          </div>
          {(selectedField && (
            <div className={styles.field}>
              <Label>Tag Selected</Label>
              <Text block nowrap={false}>
                {selectedField.name}
              </Text>
              <Label>Description</Label>
              <Text block>{selectedField.description}</Text>
              <Label>Type of Data input</Label>
              <Text block>{selectedField.dataType || "String"}</Text>
              <Label>Example Format</Label>
              <Text block>{selectedField.example || "Not Available"}</Text>
            </div>
          )) || (
            <div className={styles.intro}>
              <Text className={styles.introHeader} block>
                Welcome to FIX Marketplace Annotation Tool
              </Text>
              <Text className={styles.introHelp} block>
                Start by selecting a tag to annotate. Then hit <em>Insert Tag</em> to add the tag to your document.
              </Text>
            </div>
          )}
        </div>
        <div className={styles.footer}>
          <DefaultButton text="Clear" onClick={onClear} />
          <PrimaryButton text="Insert Tag" onClick={onInsert} disabled={!selectedField} />
        </div>
      </div>
    </>
  );
}
