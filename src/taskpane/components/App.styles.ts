import { makeStyles } from "@fluentui/react-components";

export const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    padding: "16px",
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
    marginBottom: "8px",
  },
  userMessage: {
    marginLeft: "auto",
    backgroundColor: "#e3f2fd",
    padding: "8px 12px",
    borderRadius: "18px 18px 0 18px",
    maxWidth: "80%",
  },
  botMessage: {
    marginRight: "auto",
    backgroundColor: "#f5f5f5",
    padding: "8px 12px",
    borderRadius: "18px 18px 18px 0",
    maxWidth: "80%",
  },
  inputContainer: {
    display: "flex",
    gap: "8px",
  },
  inputField: {
    flex: 1,
  },
});
