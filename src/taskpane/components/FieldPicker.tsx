import {
  Callout,
  DefaultSpacing,
  FluentTheme,
  FocusZone,
  FocusZoneDirection,
  Icon,
  KeyCodes,
  List,
  mergeStyleSets,
  SearchBox,
} from "@fluentui/react";
import FuzzySearch from "fuse.js";
import * as React from "react";
import { Field } from "./api";

const styles = mergeStyleSets({
  item: {
    padding: `${DefaultSpacing.s1}`,
    cursor: "pointer",
    selectors: {
      "&:hover, &:focus": {
        backgroundColor: FluentTheme.semanticColors.buttonBackgroundHovered,
      },
    },
  },
  placeholder: {
    padding: `${DefaultSpacing.s1} ${DefaultSpacing.m}`,
    display: "grid",
    alignItems: "center",
    gridColumnGap: 8,
    gridTemplateColumns: "min-content auto",
  },
});

interface FieldPickerProps {
  fields: Field[];
  onSelect: (field: Field) => void;
  onFilter: (filter: string) => void;
  value?: string;
}

export function FieldPicker(props: FieldPickerProps): React.ReactElement {
  const { value, fields, onSelect, onFilter } = props;
  const containerRef = React.useRef<HTMLDivElement>();
  const index = React.useMemo(
    () =>
      new FuzzySearch(fields, {
        minMatchCharLength: 0,
        keys: ["displayName"],
      }),
    [fields]
  );
  const [suggestionVisible, setSuggestionVisible] = React.useState<boolean>();
  const [suggestions, setSuggestions] = React.useState<Field[]>(fields);

  const handleSuggestionSelect = React.useCallback(
    (field: Field) => {
      setSuggestionVisible(false);
      onSelect(field);
      onFilter(field.displayName);
    },
    [onSelect]
  );

  const handleFilter = React.useCallback(
    (_, filter: string) => {
      filter = filter || "";
      let suggestions = [];
      if (!filter) {
        suggestions = index["_docs"];
      } else {
        suggestions = index.search(filter).map((e) => e.item);
      }
      if (!suggestions || !suggestions.length) {
        suggestions = [{ name: undefined } as Field];
      }
      setSuggestions(suggestions);
      setSuggestionVisible(true);
      onFilter(filter);
    },
    [index]
  );

  const handleKeyDown = React.useCallback((ev: React.KeyboardEvent<HTMLElement>): void => {
    switch (ev.keyCode) {
      case KeyCodes.down:
        let el: any = window.document.querySelector("#SearchList");
        el.focus();
        break;
    }
  }, []);

  const handleRenderCell = React.useCallback(
    (item: Field) => {
      if (!item.name) {
        return (
          <div key="dummy" className={styles.placeholder}>
            <Icon iconName="Sad"></Icon>
            <span>No tags found</span>
          </div>
        );
      }
      return (
        <div
          key={item.name}
          className={styles.item}
          data-is-focusable={true}
          onKeyDown={(ev: React.KeyboardEvent<HTMLElement>) => handleListItemKeyDown(ev, item)}
          onClick={() => handleSuggestionSelect(item)}
        >
          {item.displayName}
        </div>
      );
    },
    [onSelect]
  );

  const handleListItemKeyDown = React.useCallback(
    (ev: React.KeyboardEvent<HTMLElement>, item: Field): void => {
      const keyCode = ev.which;
      switch (keyCode) {
        case KeyCodes.enter:
          handleSuggestionSelect(item);
          break;
      }
    },
    [handleSuggestionSelect]
  );

  return (
    <div ref={containerRef} onKeyDown={handleKeyDown}>
      <SearchBox
        placeholder="Search for tag..."
        onFocus={() => setSuggestionVisible(true)}
        autoComplete="off"
        value={value}
        onChange={handleFilter}
      />
      <Callout
        gapSpace={2}
        coverTarget={false}
        alignTargetEdge={true}
        onDismiss={() => setSuggestionVisible(false)}
        hidden={!suggestionVisible}
        calloutMaxHeight={300}
        style={{ width: containerRef.current?.clientWidth, overflowY: "auto" }}
        target={containerRef.current}
        directionalHint={5}
        isBeakVisible={false}
        shouldUpdateWhenHidden={true}
      >
        <FocusZone direction={FocusZoneDirection.vertical}>
          <List id="SearchList" tabIndex={0} items={suggestions} onRenderCell={handleRenderCell} />
        </FocusZone>
      </Callout>
    </div>
  );
}
