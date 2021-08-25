import { FluentTheme, mergeStyleSets } from "@fluentui/react";

export default mergeStyleSets({
  root: {
    height: "100vh",
    display: "grid",
    gridTemplateRows: "auto 1fr auto",
  },
  intro: {
    padding: "2rem",
    maxWidth: "80%",
    background: "rgb(243, 242, 241)",
    textAlign: "center",
    opacity: ".8",
    flex: 1,
    overflow: "hidden",
    display: "flex",
    alignItems: "center",
    flexDirection: "column",
    justifyContent: "center",
    lineHeight: "1.5",
  },
  introHeader: {
    fontWeight: "bold",
    marginBottom: "0.7em",
    color: FluentTheme.palette.themePrimary,
  },
  introHelp: {},
  header: {
    padding: "4px 16px",
    minHeight: 40,
    backgroundColor: "#172733",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  body: {
    overflow: "auto",
    padding: "8px 16px 0",
    display: "flex",
    flexDirection: "column",
  },
  product: {
    margin: "-8px -16px",
    padding: "10px 16px",
    background: FluentTheme.semanticColors.buttonBackgroundDisabled,
    display: "grid",
    gridTemplateColumns: "auto min-content",
    gridColumnGap: 8,
  },
  selectedProduct: {
    backgroundColor: "rgb(176, 185, 192)",
    color: "white",
    padding: "2px 12px",
    borderRadius: 4,
    marginRight: 8,
  },
  tag: {
    margin: "12px 0 20px 0;",
  },
  field: {
    "& > label": {
      padding: "8px 12px",
      background: FluentTheme.semanticColors.buttonBackgroundDisabled,
      borderRadius: FluentTheme.effects.roundedCorner2,
    },
    "& > span": {
      padding: "8px 12px",
      marginBottom: 8,
      maxWidth: "100%",
      overflow: "auto",
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
