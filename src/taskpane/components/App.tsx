import { Label, MessageBar, MessageBarType, ResponsiveMode, Text } from "@fluentui/react";
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

export function App() {
  const [refresh, setRefresh] = React.useState(0);
  const [isLoading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string>(null);
  const [fields, setFields] = React.useState<Field[]>([]);
  const [selectedField, setSelectedField] = React.useState<Field>(null);
  const [selectedProducts, setSelectedProducts] = React.useState<string[]>([]);
  const [fieldFilter, setFieldFilter] = React.useState<string>("");
  const filteredFields = React.useMemo(() => {
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
      .getFields(refresh > 0) // true when user force-refresh schema (Refresh button)
      .then((fields) => {
        console.debug("loaded fields", fields);
        setFields(fields.sort((a, b) => a.displayName.localeCompare(b.displayName)));
      })
      .catch((err) => {
        setError(`Error loading schema. Error: ${err.message || err}`);
        console.error(err);
      })
      .finally(() => setLoading(false));
  }, [refresh]);

  React.useEffect(() => {
    if (!fields?.length) {
      return;
    }
    officeApi.registerSelectionListener((fieldName) => {
      const field = fieldName && fields.find((e) => e.displayName === fieldName);
      if (field) {
        setSelectedField(field);
        setFieldFilter(field.displayName);
      }
    });
  }, [fields]);

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

  const handleFieldSelect = React.useCallback((field) => {
    setSelectedField(field);
  }, []);

  const handleFieldFilterChange = React.useCallback((filter) => {
    setFieldFilter(filter);
  }, []);

  const handleProductSelect = React.useCallback((_, item) => {
    if (item) {
      setSelectedProducts((products) => {
        return item.selected ? [...products, item.key as string] : products.filter((key) => key !== item.key);
      });
    }
  }, []);

  const handleClear = React.useCallback(() => {
    setSelectedField(null);
    setSelectedProducts([]);
    setFieldFilter("");
  }, []);

  const handleSubmit = React.useCallback(() => {
    officeApi.insertField(selectedField);
  }, [selectedField]);

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
              onChange={handleProductSelect}
              multiSelect
              responsiveMode={ResponsiveMode.unknown}
            />
            <IconButton iconProps={{ iconName: "MoreVertical" }} menuProps={menuProps} onRenderMenuIcon={() => null} />
          </div>
          <div className={styles.tag}>
            <Label required>Select tag to annotate</Label>
            <FieldPicker
              fields={filteredFields}
              onSelect={handleFieldSelect}
              value={fieldFilter}
              onFilter={handleFieldFilterChange}
            />
          </div>
          {(selectedField && (
            <div className={styles.field}>
              <Label>Tag Selected</Label>
              <Text block nowrap={false}>
                {selectedField.displayName}
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
          <DefaultButton text="Clear" onClick={handleClear} />
          <PrimaryButton text="Insert Tag" onClick={handleSubmit} disabled={!selectedField} />
        </div>
      </div>
    </>
  );
}
