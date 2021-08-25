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
import * as React from "react";
import { Field } from "./api";

const styles = mergeStyleSets({
  root: {
    padding: `${DefaultSpacing.s1}`,
    cursor: "pointer",
    selectors: {
      "&:hover": {
        backgroundColor: FluentTheme.semanticColors.buttonBackgroundHovered,
      },
    },
  },
  placeholder: {
    padding: `${DefaultSpacing.s1} ${DefaultSpacing.m}`,
    display: 'grid',
    alignItems: 'center',
    gridColumnGap: 8,
    gridTemplateColumns: 'min-content auto',
  },
});

interface FieldPickerProps {
  fields: Field[];
  onSelect: (field: Field) => void;
  className?: string;
}

export function FieldPicker({ fields, onSelect, className }: FieldPickerProps): React.ReactElement {
  const containerRef = React.useRef<HTMLDivElement>();
  const [filter, setFilter] = React.useState<string>('');
  const [suggestionVisible, setSuggestionVisible] = React.useState<boolean>();
  const [suggestions, setSuggestions] = React.useState<Field[]>([]);

  const onSuggestionSelect = React.useCallback(
    (field: Field) => {
      setFilter(field.name);
      setSuggestionVisible(false);
      onSelect(field);
    },
    [onSelect]
  );

  const onFilter = React.useCallback(
    (_, text: string) => {
      const filter = text || '';
      setFilter(filter);
      setSuggestionVisible(true);
      let suggestions = (fields || []).filter((e) => e.name.toLowerCase().includes((filter || "").trim().toLowerCase()));
      if (!suggestions || !suggestions.length) {
        suggestions = [{ name: undefined } as Field];
      }
      setSuggestions(suggestions);
    },
    [fields]
  );

  const onKeyDown = React.useCallback((ev: React.KeyboardEvent<HTMLElement>): void => {
    switch (ev.keyCode) {
      case KeyCodes.down:
        let el: any = window.document.querySelector("#SearchList");
        el.focus();
        break;
    }
  }, []);

  const onRenderCell = React.useCallback(
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
          className={styles.root}
          data-is-focusable={true}
          onKeyDown={(ev: React.KeyboardEvent<HTMLElement>) => handleListItemKeyDown(ev, item)}
          onClick={() => onSuggestionSelect(item)}
        >
          {item.name}
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
          onSuggestionSelect(item);
          break;
      }
    },
    [onSuggestionSelect]
  );

  return (
    <div ref={containerRef} className={className} onKeyDown={onKeyDown}>
      <SearchBox
        placeholder="Search for tag..."
        onFocus={() => setSuggestionVisible(true)}
        autoComplete="off"
        value={filter}
        onChange={onFilter}
      />
      <Callout
        gapSpace={2}
        coverTarget={false}
        alignTargetEdge={true}
        onDismiss={() => setSuggestionVisible(false)}
        // setInitialFocus={true}
        hidden={!suggestionVisible}
        calloutMaxHeight={300}
        style={{ width: containerRef.current?.clientWidth, overflowY: "auto" }}
        target={containerRef.current}
        directionalHint={5}
        isBeakVisible={false}
      >
        <FocusZone direction={FocusZoneDirection.vertical}>
          <List id="SearchList" tabIndex={0} items={suggestions} onRenderCell={onRenderCell} />
        </FocusZone>
      </Callout>
    </div>
  );
}
