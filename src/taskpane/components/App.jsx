import * as React from "react";
import PropTypes from "prop-types";
import PromptConfig from "./PromptConfig";
import { Button, makeStyles, tokens, FluentProvider, teamsLightTheme, teamsDarkTheme, Switch, Label } from "@fluentui/react-components";
import { getSuggestedReply } from "../geminiApi";

const useStyles = makeStyles({
  root: {
    height: "100vh",
    background: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "8px 16px",
    boxSizing: "border-box",
  },
  headerContainer: {
    width: "100%",
    display: "flex",
    justifyContent: "flex-end",
    alignItems: "center",
    padding: "0 16px",
    boxSizing: "border-box",
    marginBottom: "16px",
  },
  buttonGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "16px",
    marginTop: "0",
    marginBottom: "12px",
    width: "100%",
    maxWidth: "600px",
  },
  contentArea: {
    width: "100%",
    minHeight: "300px",
    background: tokens.colorNeutralBackground4,
    borderRadius: tokens.borderRadiusMedium,
    padding: "24px",
    marginTop: "0",
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase400,
    boxShadow: tokens.shadow8,
    wordBreak: "break-word",
    overflowY: "auto",
    flex: 1,
    display: "flex",
    flexDirection: "column",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    boxSizing: "border-box",
    marginBottom: "8px",
  },
  activeButton: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundInverted,
    fontWeight: "bold",
    border: `2px solid ${tokens.colorBrandForeground1}`,
    boxShadow: tokens.shadow8,
    "&:hover": {
      backgroundColor: tokens.colorBrandBackgroundHover,
    },
  },
  gridButton: {
    width: "100%",
    height: "100%",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    textAlign: "center",
    whiteSpace: "nowrap",
    lineHeight: "normal",
  }
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [generatedContent, setGeneratedContent] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [templates, setTemplates] = React.useState([]);
  const [activeButton, setActiveButton] = React.useState(null);
  // const [isDarkMode, setIsDarkMode] = React.useState(false); // dark mode temporarily disabled
  const [customPrompts, setCustomPrompts] = React.useState({
    suggestReply: "Suggest a professional reply to this email:\n{emailBody}",
    personalize: "Personalize a reply to this email for a sales manager named Jamie, referencing the Q3 report:\n{emailBody}",
    summarize: "Summarize this email in 2 sentences:\n{emailBody}",
    extractActions: "Extract all action items from this email:\n{emailBody}",
    salesInsights: "Provide insights and urgency analysis for this email:\n{emailBody}",
  });

  // Responsive header style
  const headerTitle = "SalesGenie AI";
  const headerLogo = "assets/logo-filled.webp";
  

  // Helper to get email body from Outlook
  let getEmailBody = () => {
    return new Promise((resolve, reject) => {
      if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        Office.context.mailbox.item.body.getAsync("text", (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject("Failed to get email body.");
          }
        });
      } else {
        reject("Office.js not available or not in Outlook context.");
      }
    });
  };
  
  // getEmailBody= () => "I am a sales manager and I am sending this email to you";
  // Helper to call Gemini API with a custom prompt
  const callGemini = async (promptTemplate) => {
    setLoading(true);
    setGeneratedContent("Generating...");
    try {
      const emailBody = await getEmailBody();
      const prompt = promptTemplate.replace("{emailBody}", emailBody);
      console.log("prompt", prompt);
      const reply = await getSuggestedReply(prompt);
      setGeneratedContent(reply);
    } catch (e) {
      setGeneratedContent("Error: " + e);
    }
    setLoading(false);
  };

  // Feature handlers
  const handleSuggestReply = () => {
    setActiveButton('suggestReply');
    callGemini(customPrompts.suggestReply);
  };
  const handlePersonalize = () => {
    setActiveButton('personalize');
    callGemini(customPrompts.personalize);
  };
  const handleSummarize = () => {
    setActiveButton('summarize');
    callGemini(customPrompts.summarize);
  };
  const handleExtractActions = () => {
    setActiveButton('extractActions');
    callGemini(customPrompts.extractActions);
  };
  const handleSalesInsights = () => {
    setActiveButton('salesInsights');
    callGemini(customPrompts.salesInsights);
  };
  const handleSaveTemplate = () => {
    setActiveButton('saveTemplate');
    setTemplates((prev) => [...prev, generatedContent]);
    setGeneratedContent("Template saved.");
  };
  const handleViewTemplates = () => {
    setActiveButton('viewTemplates');
    if (templates.length === 0) {
      setGeneratedContent("No templates saved.");
    } else {
      setGeneratedContent(templates.map((t, i) => `Template ${i + 1}:\n${t}`).join("\n\n---\n\n"));
    }
  };
  const handleClear = () => {
    setActiveButton('clear');
    setGeneratedContent("");
  };

  const handleSavePrompts = (newPrompts) => {
    setCustomPrompts(newPrompts);
  };

  // const toggleDarkMode = () => {
  //   setIsDarkMode(!isDarkMode);
  // };

  return (
    <FluentProvider theme={teamsLightTheme}> {/* dark mode disabled for now */}
      <div className={styles.root}>
        <div className={styles.buttonGrid}>
          <Button 
            appearance={activeButton === 'suggestReply' ? "primary" : "secondary"}
            onClick={handleSuggestReply} 
            disabled={loading}
            className={activeButton === 'suggestReply' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Suggest Reply
          </Button>
          {/*
          <Button 
            appearance={activeButton === 'personalize' ? "primary" : "secondary"}
            onClick={handlePersonalize} 
            disabled={loading}
            className={activeButton === 'personalize' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Personalize
          </Button>
          */}
          <Button 
            appearance={activeButton === 'summarize' ? "primary" : "secondary"}
            onClick={handleSummarize} 
            disabled={loading}
            className={activeButton === 'summarize' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Summarize
          </Button>
          {/*
          <Button 
            appearance={activeButton === 'extractActions' ? "primary" : "secondary"}
            onClick={handleExtractActions} 
            disabled={loading}
            className={activeButton === 'extractActions' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Extract Actions
          </Button>
          */}
          <Button 
            appearance={activeButton === 'salesInsights' ? "primary" : "secondary"}
            onClick={handleSalesInsights} 
            disabled={loading}
            className={activeButton === 'salesInsights' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
          Insights
          </Button>
          {/*
          <Button 
            appearance={activeButton === 'saveTemplate' ? "primary" : "secondary"}
            onClick={handleSaveTemplate} 
            disabled={loading || !generatedContent}
            className={activeButton === 'saveTemplate' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Save Template
          </Button>
          */}
          {/*
          <Button 
            appearance={activeButton === 'viewTemplates' ? "primary" : "secondary"}
            onClick={handleViewTemplates} 
            disabled={loading}
            className={activeButton === 'viewTemplates' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            View Templates
          </Button>
          */}
          <div className={styles.gridButton} style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', whiteSpace: 'nowrap' }}>
            <PromptConfig onSavePrompts={handleSavePrompts} />
          </div>
          {/* Dark mode toggle disabled for now */}
        </div>
        <div
          className={styles.contentArea}
          dangerouslySetInnerHTML={{
            __html: generatedContent
              ? generatedContent
                  .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener noreferrer">$1</a>') // links clickable
                  .replace(/\n/g, '<br>') // preserve line breaks
                  .replace(/\*\*(.*?)\*\*/g, '<b>$1</b>') // bold for **text**
                  .replace(/\*(.*?)\*/g, '<i>$1</i>') // italics for *text*
              : "Generated content will appear here."
          }}
        />
      </div>
    </FluentProvider>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;

