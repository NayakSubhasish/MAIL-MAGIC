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
    writeEmail: "Write a professional email based on this description: {emailBody}",
    editEmail: "Edit and improve this email for clarity and professionalism: {emailBody}",
    respondToEmail: "Generate a professional response to this email: {emailBody}",
    rewriteEmail: "Rewrite this email to be more professional and clear: {emailBody}",
    cleanUpEmail: "Clean up and improve the grammar and structure of this email: {emailBody}",
    // salesInsights: "Provide insights and urgency analysis for this email:\n{emailBody}",
  });

  // Responsive header style
  const headerTitle = "SalesGenie AI";
  const headerLogo = "assets/logo-filled.webp";
  

  // Helper to get email body from Outlook
  let getEmailBody = () => {
    return new Promise((resolve, reject) => {
      if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;
        
        // Check if we're in compose mode or read mode
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
          // For reading emails, get the current email body
          item.body.getAsync("text", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value);
            } else {
              reject("Failed to get email body.");
            }
          });
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
          // For appointments, get the body
          item.body.getAsync("text", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value);
            } else {
              reject("Failed to get appointment body.");
            }
          });
        } else {
          reject("Unsupported item type.");
        }
      } else {
        reject("Office.js not available or not in Outlook context.");
      }
    });
  };

  // Helper to get conversation thread (previous emails)
  let getConversationThread = () => {
    return new Promise((resolve, reject) => {
      if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;
        
        // Get conversation thread if available
        if (item.conversationId) {
          // For now, we'll use the current item's subject to identify the thread
          // In a full implementation, you'd query the conversation
          resolve(`Thread: ${item.subject || 'No subject'}\n\nCurrent email content will be processed.`);
        } else {
          resolve("No conversation thread available.");
        }
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
      const conversationThread = await getConversationThread();
      
      // Include conversation thread context in the prompt
      const contextWithThread = `Conversation Context:\n${conversationThread}\n\nCurrent Email:\n${emailBody}`;
      const prompt = promptTemplate.replace("{emailBody}", contextWithThread);
      
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
  // const handleSalesInsights = () => {
  //   setActiveButton('salesInsights');
  //   callGemini(customPrompts.salesInsights);
  // };
  const handleWriteEmail = () => {
    setActiveButton('writeEmail');
    callGemini(customPrompts.writeEmail);
  };
  const handleEditEmail = () => {
    setActiveButton('editEmail');
    callGemini(customPrompts.editEmail);
  };
  const handleRespondToEmail = () => {
    setActiveButton('respondToEmail');
    callGemini(customPrompts.respondToEmail);
  };
  const handleRewriteEmail = () => {
    setActiveButton('rewriteEmail');
    callGemini(customPrompts.rewriteEmail);
  };
  const handleCleanUpEmail = () => {
    setActiveButton('cleanUpEmail');
    callGemini(customPrompts.cleanUpEmail);
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
            appearance={activeButton === 'writeEmail' ? "primary" : "secondary"}
            onClick={handleWriteEmail} 
            disabled={loading}
            className={activeButton === 'writeEmail' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Write Email
          </Button>
          <Button 
            appearance={activeButton === 'editEmail' ? "primary" : "secondary"}
            onClick={handleEditEmail} 
            disabled={loading}
            className={activeButton === 'editEmail' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Edit Email
          </Button>
          <Button 
            appearance={activeButton === 'respondToEmail' ? "primary" : "secondary"}
            onClick={handleRespondToEmail} 
            disabled={loading}
            className={activeButton === 'respondToEmail' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Respond to Email
          </Button>
          <Button 
            appearance={activeButton === 'rewriteEmail' ? "primary" : "secondary"}
            onClick={handleRewriteEmail} 
            disabled={loading}
            className={activeButton === 'rewriteEmail' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Rewrite Email
          </Button>
          <Button 
            appearance={activeButton === 'cleanUpEmail' ? "primary" : "secondary"}
            onClick={handleCleanUpEmail} 
            disabled={loading}
            className={activeButton === 'cleanUpEmail' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Clean-Up Email
          </Button>
          <Button 
            appearance={activeButton === 'summarize' ? "primary" : "secondary"}
            onClick={handleSummarize} 
            disabled={loading}
            className={activeButton === 'summarize' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Summarize
          </Button>
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

