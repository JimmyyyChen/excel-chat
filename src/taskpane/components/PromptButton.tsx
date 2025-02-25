import * as React from "react";
import { Button } from "@fluentui/react-components";

interface PromptButtonProps {
  text: string;
  promptText: string;
  onClick: (promptText: string) => void;
}

export const PromptButton: React.FC<PromptButtonProps> = ({ text, promptText, onClick }) => {
  return (
    <Button appearance="outline" size="small" onClick={() => onClick(promptText)}>
      {text}
    </Button>
  );
}; 