import {
  IComboBoxOption,
  Label,
  MessageBar,
  MessageBarType,
  ResponsiveMode,
  Text,
  VirtualizedComboBox,
} from "@fluentui/react";
import { DefaultButton, IconButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import * as React from "react";
import { createUseStyles } from "react-jss";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import officeApi, { Field } from "./api";

const useStyles = createUseStyles({
  root: {
    height: "100vh",
    display: "grid",
    gridTemplateRows: "auto 1fr auto",
  },
  header: {
    padding: "4px 16px",
    backgroundColor: "#172733",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  body: {
    overflow: "auto",
    padding: "8px 16px 0",
  },
  tag: {
    margin: "12px 0 20px 0;",
  },
  field: {
    "& > label": {
      padding: "8px 12px",
      background: "rgb(243, 242, 241)",
      borderRadius: 2,
      "& + *": {
        padding: "8px 12px",
        marginBottom: 16,
        maxWidth: "100%",
        overflow: "auto",
      },
    },
  },
  footer: {
    display: "grid",
    padding: "16px 16px",
    gridColumnGap: 16,
    gridTemplateColumns: "1fr 1fr",
  },
  center: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    height: "100vh",
    backgroundColor: "white",
  },
});

const products: IDropdownOption[] = [
  { key: "CD", text: "CD" },
  { key: "CP", text: "CP" },
  { key: "MTN", text: "MTN" },
];

interface FieldSelection extends IComboBoxOption, Field {}

export function App() {
  const classes = useStyles();
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
      return
    }
    const field = fields.find(e => e.name === selectedFieldId);
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

  const onInsert = React.useCallback(() => {
    officeApi.insertField(selectedField);
  }, [selectedField]);

  const onSelectField = React.useCallback((_, option) => {
    setSelectedField(option as FieldSelection);
  }, []);

  const onSelectProduct = React.useCallback((_, item) => {
    if (item) {
      setSelectedProducts(
        item.selected ? [...selectedProducts, item.key as string] : selectedProducts.filter((key) => key !== item.key)
      );
    }
  }, []);

  if (isLoading) {
    return (
      <div className={classes.center}>
        <Spinner label="Loading schema..." size={SpinnerSize.large} />
      </div>
    );
  }

  if (error) {
    return (
      <div className={classes.center}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {error}
        </MessageBar>
      </div>
    );
  }

  return (
    <>
      <div className={classes.root}>
        <div className={classes.header}>
          <img src="assets/logo_inversed.svg" />
          <IconButton iconProps={{ iconName: "MoreVertical" }} menuProps={menuProps} onRenderMenuIcon={() => null} />
        </div>
        <div className={classes.body}>
          <Dropdown
            label="Products"
            placeholder="Select products"
            options={products}
            onChange={onSelectProduct}
            multiSelect
            responsiveMode={ResponsiveMode.unknown}
          />
          <VirtualizedComboBox
            className={classes.tag}
            required
            label="Select tag to annotate"
            placeholder="Search for tag..."
            selectedKey={selectedField?.key || null}
            onChange={onSelectField}
            allowFreeform
            scrollSelectedToTop
            autoComplete="on"
            options={filteredFields}
            useComboBoxAsMenuWidth
          />
          {selectedField && (
            <div className={classes.field}>
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
          )}
        </div>
        <div className={classes.footer}>
          <DefaultButton text="Clear" onClick={() => setSelectedField(null)} />
          <PrimaryButton text="Insert Tag" onClick={onInsert} disabled={!selectedField} />
        </div>
      </div>
    </>
  );
}
