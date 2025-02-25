import { makeStyles } from "@fluentui/react-components";

export const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    padding: "16px",
    boxSizing: "border-box",
  },
  chatContainer: {
    flex: 1,
    overflowY: "auto",
    marginBottom: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  messageRow: {
    display: "flex",
    width: "100%",
  },
  userMessage: {
    alignSelf: "flex-end",
    backgroundColor: "#e6f7ff",
    marginLeft: "auto",
    maxWidth: "80%",
  },
  botMessage: {
    alignSelf: "flex-start",
    backgroundColor: "#f0f0f0",
    marginRight: "auto",
    maxWidth: "80%",
  },
  inputContainer: {
    display: "flex",
    gap: "8px",
  },
  inputField: {
    flex: 1,
  },
  recommendedPrompts: {
    display: "flex",
    gap: "8px",
    marginBottom: "16px",
    flexWrap: "wrap",
  },
  messageImage: {
    maxWidth: "100%",
    maxHeight: "200px",
  },
  messageTable: {
    borderCollapse: "collapse",
    width: "100%",
    "& th, & td": {
      border: "1px solid #ddd",
      padding: "8px",
      textAlign: "left",
    },
    "& th": {
      backgroundColor: "#f2f2f2",
      fontWeight: "bold",
    },
    "& tr:nth-child(even)": {
      backgroundColor: "#f9f9f9",
    },
  },
  tableActions: {
    marginTop: "8px",
    display: "flex",
    justifyContent: "flex-end",
  },
});
