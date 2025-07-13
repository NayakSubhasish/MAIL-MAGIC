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
    gap: "12px",
    marginTop: "0",
    marginBottom: "16px",
    width: "100%",
    maxWidth: "600px",
    padding: "0 8px",
  },
  contentArea: {
    width: "100%",
    minHeight: "300px",
    background: "#ffffff",
    borderRadius: "8px",
    padding: "20px",
    marginTop: "8px",
    marginLeft: "8px",
    marginRight: "8px",
    color: "#323130",
    fontSize: "14px",
    lineHeight: "1.5",
    boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
    wordBreak: "break-word",
    overflowY: "auto",
    flex: 1,
    display: "flex",
    flexDirection: "column",
    border: "1px solid #e1e1e1",
  },
  gridButton: {
    minHeight: "48px",
    fontSize: "14px",
    fontWeight: "600",
    borderRadius: "8px",
    border: "1px solid #d1d1d1",
    backgroundColor: "#ffffff",
    color: "#323130",
    transition: "all 0.2s ease",
    cursor: "pointer",
    "&:hover": {
      backgroundColor: "#f3f2f1",
      borderColor: "#0078d4",
    },
    "&:active": {
      backgroundColor: "#edebe9",
    },
  },
  activeButton: {
    backgroundColor: "#0078d4",
    borderColor: "#0078d4",
    color: "#ffffff",
    "&:hover": {
      backgroundColor: "#106ebe",
      borderColor: "#106ebe",
    },
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [generatedContent, setGeneratedContent] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [templates, setTemplates] = React.useState([]);
  const [activeButton, setActiveButton] = React.useState(null);
  const [showWriteEmailForm, setShowWriteEmailForm] = React.useState(false);
  const [emailForm, setEmailForm] = React.useState({
    description: "",
    additionalInstructions: "",
    tone: "Formal",
    pointOfView: "Organization perspective"
  });
  // const [isDarkMode, setIsDarkMode] = React.useState(false); // dark mode temporarily disabled
  const [customPrompts, setCustomPrompts] = React.useState({
    suggestReply: "Suggest a professional reply to this email:\n{emailBody}",
    summarize: "Summarize this email in 2 sentences:\n{emailBody}",
    writeEmail: "Write a professional email with the following details:\nDescription: {description}\nAdditional Instructions: {additionalInstructions}\nTone: {tone}\nPoint of View: {pointOfView}",
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
    setShowWriteEmailForm(false);
    callGemini(customPrompts.suggestReply);
  };
  const handleSummarize = () => {
    setActiveButton('summarize');
    setShowWriteEmailForm(false);
    callGemini(customPrompts.summarize);
  };
  const handleWriteEmail = () => {
    setActiveButton('writeEmail');
    setShowWriteEmailForm(true);
    setGeneratedContent("");
  };
  const handleGenerateEmail = () => {
    if (!emailForm.description.trim()) {
      setGeneratedContent("Please enter a description for the email.");
      return;
    }
    
    const prompt = customPrompts.writeEmail
      .replace("{description}", emailForm.description)
      .replace("{additionalInstructions}", emailForm.additionalInstructions || "None")
      .replace("{tone}", emailForm.tone)
      .replace("{pointOfView}", emailForm.pointOfView);
    
    setLoading(true);
    setGeneratedContent("Generating email...");
    getSuggestedReply(prompt).then(reply => {
      setGeneratedContent(reply);
      setLoading(false);
    }).catch(e => {
      setGeneratedContent("Error: " + e);
      setLoading(false);
    });
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
            appearance={activeButton === 'suggestReply' ? "primary" : "secondary"}
            onClick={handleSuggestReply} 
            disabled={loading}
            className={activeButton === 'suggestReply' ? `${styles.activeButton} ${styles.gridButton}` : styles.gridButton}
          >
            Suggest Reply
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
        
        {showWriteEmailForm && (
          <div style={{ 
            padding: '20px', 
            borderTop: '1px solid #e1e1e1', 
            marginTop: '16px',
            backgroundColor: '#fafafa',
            borderRadius: '8px',
            marginLeft: '8px',
            marginRight: '8px'
          }}>
            <h3 style={{ 
              margin: '0 0 20px 0', 
              fontSize: '18px', 
              fontWeight: '600',
              color: '#323130',
              textAlign: 'center'
            }}>Write New Email</h3>
            
            <div style={{ marginBottom: '20px' }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '8px', 
                fontWeight: '600',
                fontSize: '14px',
                color: '#323130'
              }}>Description *</label>
              <textarea
                placeholder="Describe what you want to write about and who the email is for."
                value={emailForm.description}
                onChange={(e) => setEmailForm({...emailForm, description: e.target.value})}
                style={{
                  width: '100%',
                  minHeight: '90px',
                  padding: '12px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '6px',
                  fontSize: '14px',
                  fontFamily: 'inherit',
                  resize: 'vertical',
                  boxSizing: 'border-box',
                  backgroundColor: '#ffffff',
                  transition: 'border-color 0.2s ease',
                  outline: 'none'
                }}
                onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
              />
            </div>
            
            <div style={{ marginBottom: '20px' }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '8px', 
                fontWeight: '600',
                fontSize: '14px',
                color: '#323130'
              }}>Additional Instructions</label>
              <textarea
                placeholder="Any additional instructions or specific requirements..."
                value={emailForm.additionalInstructions}
                onChange={(e) => setEmailForm({...emailForm, additionalInstructions: e.target.value})}
                style={{
                  width: '100%',
                  minHeight: '70px',
                  padding: '12px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '6px',
                  fontSize: '14px',
                  fontFamily: 'inherit',
                  resize: 'vertical',
                  boxSizing: 'border-box',
                  backgroundColor: '#ffffff',
                  transition: 'border-color 0.2s ease',
                  outline: 'none'
                }}
                onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
              />
            </div>
            
            <div style={{ 
              display: 'grid', 
              gridTemplateColumns: '1fr 1fr', 
              gap: '16px', 
              marginBottom: '24px' 
            }}>
              <div>
                <label style={{ 
                  display: 'block', 
                  marginBottom: '8px', 
                  fontWeight: '600',
                  fontSize: '14px',
                  color: '#323130'
                }}>Tone *</label>
                <select
                  value={emailForm.tone}
                  onChange={(e) => setEmailForm({...emailForm, tone: e.target.value})}
                  style={{
                    width: '100%',
                    padding: '12px',
                    border: '1px solid #d1d1d1',
                    borderRadius: '6px',
                    fontSize: '14px',
                    fontFamily: 'inherit',
                    backgroundColor: '#ffffff',
                    boxSizing: 'border-box',
                    outline: 'none',
                    cursor: 'pointer'
                  }}
                  onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                  onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
                >
                  <option value="Formal">Formal</option>
                  <option value="Casual">Casual</option>
                  <option value="Friendly">Friendly</option>
                  <option value="Professional">Professional</option>
                  <option value="Persuasive">Persuasive</option>
                </select>
              </div>
              
              <div>
                <label style={{ 
                  display: 'block', 
                  marginBottom: '8px', 
                  fontWeight: '600',
                  fontSize: '14px',
                  color: '#323130'
                }}>Point of View *</label>
                <select
                  value={emailForm.pointOfView}
                  onChange={(e) => setEmailForm({...emailForm, pointOfView: e.target.value})}
                  style={{
                    width: '100%',
                    padding: '12px',
                    border: '1px solid #d1d1d1',
                    borderRadius: '6px',
                    fontSize: '14px',
                    fontFamily: 'inherit',
                    backgroundColor: '#ffffff',
                    boxSizing: 'border-box',
                    outline: 'none',
                    cursor: 'pointer'
                  }}
                  onFocus={(e) => e.target.style.borderColor = '#0078d4'}
                  onBlur={(e) => e.target.style.borderColor = '#d1d1d1'}
                >
                  <option value="Organization perspective">Organization perspective</option>
                  <option value="Personal perspective">Personal perspective</option>
                  <option value="Team perspective">Team perspective</option>
                  <option value="Customer perspective">Customer perspective</option>
                </select>
              </div>
            </div>
            
            <Button
              appearance="primary"
              onClick={handleGenerateEmail}
              disabled={loading || !emailForm.description.trim()}
              style={{ 
                width: '100%',
                padding: '12px 24px',
                fontSize: '16px',
                fontWeight: '600',
                borderRadius: '6px',
                minHeight: '44px'
              }}
            >
              {loading ? 'Generating...' : 'Generate Email'}
            </Button>
          </div>
        )}
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

