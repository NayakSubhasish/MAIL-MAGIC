// DEPRECATED: This component is no longer used in the new SalesReplyAI Plugin UI.
// Safe to delete if not needed for reference.
import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import PropTypes from "prop-types";

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
});

const TextInsertion = (props) => {
  const [text, setText] = useState("Some text.");

  const handleTextInsertion = async () => {
    await props.insertText(text);
  };

  const handleTextChange = async (event) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Enter text to be inserted into the document.">
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      <Field className={styles.instructions}>Click the button to insert text.</Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextInsertion}>
        Insert text
      </Button>
    </div>
  );
};

TextInsertion.propTypes = {
  insertText: PropTypes.func.isRequired,
};

export default TextInsertion;
