import * as React from "react";
import PropTypes from "prop-types";
import PromptConfig from "./PromptConfig";
import { Button, makeStyles, tokens, FluentProvider, teamsLightTheme, teamsDarkTheme, Switch, Label, Tab, TabList } from "@fluentui/react-components";
import { getSuggestedReply } from "../botAtWorkApi";

const useStyles = makeStyles({
  root: {
    height: "100vh",
    background: "linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "16px",
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
  tabContainer: {
    width: "100%",
    maxWidth: "400px",
    marginBottom: "10px",
  },
  contentArea: {
    width: "100%",
    minHeight: "300px",
    background: "linear-gradient(145deg, #e8e8e8, #d4d4d4)",
    borderRadius: "8px",
    padding: "16px",
    color: "#2d2d2d",
    fontSize: "15px",
    lineHeight: "1.7",
    boxShadow: "0 4px 20px rgba(0,0,0,0.12), 0 1px 3px rgba(0,0,0,0.15)",
    wordBreak: "break-word",
    overflowY: "auto",
    flex: 1,
    display: "flex",
    flexDirection: "column",
    border: "1px solid rgba(0,0,0,0.12)",
    margin: "4px 8px 4px 8px",
    transition: "all 0.3s ease",
    maxWidth: "none",
    "&:hover": {
      boxShadow: "0 8px 30px rgba(0,0,0,0.18), 0 2px 6px rgba(0,0,0,0.15)",
    },
  },
  gridButton: {
    minHeight: "44px",
    fontSize: "14px",
    fontWeight: "600",
    borderRadius: "4px",
    border: "1px solid rgba(0,120,212,0.2)",
    backgroundColor: "rgba(255,255,255,0.9)",
    color: "#323130",
    transition: "all 0.3s ease",
    cursor: "pointer",
    backdropFilter: "blur(10px)",
    boxShadow: "0 2px 8px rgba(0,0,0,0.05)",
    "&:hover": {
      backgroundColor: "rgba(255,255,255,1)",
      borderColor: "rgba(0,120,212,0.4)",
      transform: "translateY(-1px)",
      boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
    },
    "&:active": {
      transform: "translateY(0px)",
    },
  },
  activeButton: {
    background: "linear-gradient(135deg, #0078d4, #106ebe)",
    borderColor: "rgba(0,120,212,0.3)",
    color: "#ffffff",
    boxShadow: "0 4px 15px rgba(0,120,212,0.3)",
    "&:hover": {
      background: "linear-gradient(135deg, #106ebe, #005a9e)",
      transform: "translateY(-1px)",
      boxShadow: "0 6px 20px rgba(0,120,212,0.4)",
    },
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [generatedContent, setGeneratedContent] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [templates, setTemplates] = React.useState([]);
  const [activeTab, setActiveTab] = React.useState("writeEmail");
  const [showWriteEmailForm, setShowWriteEmailForm] = React.useState(true);
  const [emailForm, setEmailForm] = React.useState({
    description: "",
    // additionalInstructions: "", // Commented out as per user request
    tone: "Formal",
    pointOfView: "Organization perspective"
  });
  // const [isDarkMode, setIsDarkMode] = React.useState(false); // dark mode temporarily disabled
  const [customPrompts, setCustomPrompts] = React.useState({
    suggestReply: "Suggest a professional reply to this email:\n{emailBody}",
    summarize: "Summarize this email in 2 sentences:\n{emailBody}",
    writeEmail: "Write a professional email with the following details:\nDescription: {description}\nTone: {tone}\nPoint of View: {pointOfView}",
  });
  const [chatInput, setChatInput] = React.useState("");
  const [chatHistory, setChatHistory] = React.useState([]);
  const [isFirstResponse, setIsFirstResponse] = React.useState(true);

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
      
      // Monitor for retry attempts
      let retryCount = 0;
      const originalLog = console.log;
      console.log = (...args) => {
        if (args[0] && args[0].includes && args[0].includes('BotAtWork API attempt')) {
          retryCount++;
          if (retryCount > 1) {
            setGeneratedContent(`Generating... (API busy, retrying - attempt ${retryCount})`);
          }
        }
        originalLog.apply(console, args);
      };
      
      const reply = await getSuggestedReply(prompt);
      
      // Restore original console.log
      console.log = originalLog;
      
      setGeneratedContent(reply);
    } catch (e) {
      setGeneratedContent("Error: " + e);
    }
    setLoading(false);
  };

  // Chat input send handler: send direct prompt to LLM
  const handleChatSend = async () => {
    if (!chatInput.trim()) return;
    
    const userMessage = chatInput.trim();
    setChatInput("");
    
    // Add user message to chat history
    const newUserMessage = { type: 'user', content: userMessage };
    setChatHistory(prev => [...prev, newUserMessage]);
    
    setLoading(true);
    
    try {
      let prompt;
      if (isFirstResponse && generatedContent && generatedContent !== "Generating...") {
        // If there's existing content, include it in context
        prompt = `Based on this previous response: "${generatedContent}"\n\nUser request: ${userMessage}`;
        setIsFirstResponse(false);
      } else {
        // Build context from chat history
        const context = chatHistory.map(msg => 
          msg.type === 'user' ? `User: ${msg.content}` : `Assistant: ${msg.content}`
        ).join('\n');
        prompt = context ? `${context}\nUser: ${userMessage}` : userMessage;
      }
      
      // Monitor for retry attempts
      let retryCount = 0;
      const originalLog = console.log;
      console.log = (...args) => {
        if (args[0] && args[0].includes && args[0].includes('BotAtWork API attempt')) {
          retryCount++;
          if (retryCount > 1) {
            setGeneratedContent(`Processing... (API busy, retrying - attempt ${retryCount})`);
          }
        }
        originalLog.apply(console, args);
      };
      
      const reply = await getSuggestedReply(prompt);
      
      // Restore original console.log
      console.log = originalLog;
      
      // Add AI response to chat history
      const newAIMessage = { type: 'ai', content: reply };
      setChatHistory(prev => [...prev, newAIMessage]);
      
      // Update the main content area with latest response
      setGeneratedContent(reply);
    } catch (e) {
      const errorMessage = "Error: " + e;
      setChatHistory(prev => [...prev, { type: 'ai', content: errorMessage }]);
      setGeneratedContent(errorMessage);
    }
    setLoading(false);
  };

  // Tab handler
  const handleTabSelect = (event, data) => {
    setActiveTab(data.value);
    if (data.value === 'writeEmail') {
    setShowWriteEmailForm(true);
      setChatHistory([]);
      setIsFirstResponse(true);
    setGeneratedContent("");
    } else if (data.value === 'suggestReply') {
      setShowWriteEmailForm(false);
      setChatHistory([]);
      setIsFirstResponse(true);
      callGemini(customPrompts.suggestReply);
    }
  };
  const handleGenerateEmail = async () => {
    if (!emailForm.description.trim()) {
      setGeneratedContent("Please enter a description for the email.");
      return;
    }
    
    const prompt = customPrompts.writeEmail
      .replace("{description}", emailForm.description)
      // .replace("{additionalInstructions}", emailForm.additionalInstructions || "None") // Commented out as per user request
      .replace("{tone}", emailForm.tone)
      .replace("{pointOfView}", emailForm.pointOfView);
    
    setLoading(true);
    setGeneratedContent("Generating email...");
    
    try {
      // Monitor for retry attempts by intercepting console logs
      let retryCount = 0;
      const originalLog = console.log;
      console.log = (...args) => {
        if (args[0] && args[0].includes && args[0].includes('BotAtWork API attempt')) {
          retryCount++;
          if (retryCount > 1) {
            setGeneratedContent(`Generating email... (API busy, retrying - attempt ${retryCount})`);
          }
        }
        originalLog.apply(console, args);
      };
      
      const reply = await getSuggestedReply(prompt);
      
      // Restore original console.log
      console.log = originalLog;
      
      setGeneratedContent(reply);
      setLoading(false);
    } catch (e) {
      setGeneratedContent("Error: " + e);
      setLoading(false);
    }
  };
  const handleSaveTemplate = () => {
    setTemplates((prev) => [...prev, generatedContent]);
    setGeneratedContent("Template saved.");
  };
  const handleViewTemplates = () => {
    if (templates.length === 0) {
      setGeneratedContent("No templates saved.");
    } else {
      setGeneratedContent(templates.map((t, i) => `Template ${i + 1}:\n${t}`).join("\n\n---\n\n"));
    }
  };
  const handleClear = () => {
    setGeneratedContent("");
  };

  // const handleSavePrompts = (newPrompts) => {
  //   setCustomPrompts(newPrompts);
  // };

  // const toggleDarkMode = () => {
  //   setIsDarkMode(!isDarkMode);
  // };

  return (
    <FluentProvider theme={teamsLightTheme}> {/* dark mode disabled for now */}
      <div className={styles.root}>
        <div className={styles.tabContainer}>
          <TabList selectedValue={activeTab} onTabSelect={handleTabSelect}>
            <Tab value="writeEmail">Write Email</Tab>
            <Tab value="suggestReply">Suggest Reply</Tab>
          </TabList>
        </div>
        
        {showWriteEmailForm && (
          <div style={{ 
            padding: '8px', 
            borderTop: '1px solid #e1e1e1', 
            marginTop: '4px',
            backgroundColor: '#fafafa',
            borderRadius: '0',
            width: '100%',
            boxSizing: 'border-box'
          }}>
            <div style={{ marginBottom: '8px' }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '3px', 
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
                  minHeight: '60px',
                  padding: '6px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
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
            
            {/* Additional Instructions section - Hidden as per user request */}
            {/* <div style={{ marginBottom: '16px' }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '6px', 
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
                  minHeight: '60px',
                  padding: '10px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
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
            </div> */}
            
            <div style={{ marginBottom: '8px' }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '3px', 
                fontWeight: '600',
                fontSize: '14px',
                color: '#323130'
              }}>Tone *</label>
              <select
                value={emailForm.tone}
                onChange={(e) => setEmailForm({...emailForm, tone: e.target.value})}
                style={{
                  width: '100%',
                  padding: '6px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
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
                <option value="Professional">Professional</option>
                <option value="Empathetic">Empathetic</option>
              </select>
            </div>
            
            <div style={{ marginBottom: '10px' }}>
              <label style={{ 
                display: 'block', 
                marginBottom: '3px', 
                fontWeight: '600',
                fontSize: '14px',
                color: '#323130'
              }}>Point of View *</label>
              <select
                value={emailForm.pointOfView}
                onChange={(e) => setEmailForm({...emailForm, pointOfView: e.target.value})}
                style={{
                  width: '100%',
                  padding: '6px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
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
                <option value="Individual perspective">Individual perspective</option>
              </select>
            </div>
            
            <Button
              appearance={emailForm.description.trim() ? "primary" : "secondary"}
              onClick={handleGenerateEmail}
              disabled={loading || !emailForm.description.trim()}
              style={{ 
                width: '100%',
                padding: '8px 16px',
                fontSize: '15px',
                fontWeight: '600',
                borderRadius: '4px',
                minHeight: '36px',
                backgroundColor: emailForm.description.trim() ? '#0078d4' : '#f3f2f1',
                color: emailForm.description.trim() ? '#ffffff' : '#323130',
                border: emailForm.description.trim() ? 'none' : '1px solid #d1d1d1'
              }}
            >
              {loading ? 'Generating...' : 'Generate Email'}
            </Button>
          </div>
        )}
        {activeTab === 'suggestReply' && (
          <div style={{
            padding: '12px 16px',
            borderTop: '1px solid #e1e1e1',
            marginTop: '8px',
            backgroundColor: '#f0f0f0',
            borderRadius: '4px',
            width: '100%',
            boxSizing: 'border-box'
          }}>
            {/* Additional Instructions section - Hidden as per user request */}
            {/* <div style={{ marginBottom: '12px' }}>
              <label style={{ display: 'block', marginBottom: '6px', fontWeight: '600', fontSize: '14px', color: '#323130' }}>
                Additional Instructions
              </label>
              <textarea
                placeholder="Any additional instructions..."
                value={emailForm.additionalInstructions}
                onChange={(e) => setEmailForm({ ...emailForm, additionalInstructions: e.target.value })}
                style={{
                  width: '100%',
                  minHeight: '60px',
                  padding: '10px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  backgroundColor: '#ffffff',
                  boxSizing: 'border-box'
                }}
              />
            </div> */}
            <div style={{ marginBottom: '12px' }}>
              <label style={{ display: 'block', marginBottom: '6px', fontWeight: '600', fontSize: '14px', color: '#323130' }}>
                Tone
              </label>
              <select
                value={emailForm.tone}
                onChange={(e) => setEmailForm({ ...emailForm, tone: e.target.value })}
                style={{
                  width: '100%',
                  padding: '10px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  backgroundColor: '#ffffff',
                  boxSizing: 'border-box',
                  cursor: 'pointer'
                }}
              >
                <option value="Formal">Formal</option>
                <option value="Casual">Casual</option>
                <option value="Professional">Professional</option>
                <option value="Empathetic">Empathetic</option>
              </select>
            </div>
            <div style={{ marginBottom: '12px' }}>
              <label style={{ display: 'block', marginBottom: '6px', fontWeight: '600', fontSize: '14px', color: '#323130' }}>
                Point of View
              </label>
              <select
                value={emailForm.pointOfView}
                onChange={(e) => setEmailForm({ ...emailForm, pointOfView: e.target.value })}
                style={{
                  width: '100%',
                  padding: '10px',
                  border: '1px solid #d1d1d1',
                  borderRadius: '4px',
                  fontSize: '14px',
                  backgroundColor: '#ffffff',
                  boxSizing: 'border-box',
                  cursor: 'pointer'
                }}
              >
                <option value="Organization perspective">Organization perspective</option>
                <option value="Individual perspective">Individual perspective</option>
              </select>
            </div>
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
              : '<div style="display: flex; align-items: center; justify-content: center; height: 100%; color: #605e5c; font-style: italic; text-align: center; padding: 30px;"><div><div style="font-size: 16px; margin-bottom: 6px;">✨ Your generated content will appear here</div><div style="font-size: 12px; opacity: 0.8;">Click a button above to get started</div></div></div>'
          }}
        />
        {(activeTab === 'suggestReply' || (activeTab === 'writeEmail' && generatedContent && generatedContent !== "Generating..." && generatedContent !== "Generating email...")) && (
          <div style={{
            display: 'flex',
            borderTop: '1px solid #e1e1e1',
            padding: '6px',
            alignItems: 'center',
            margin: '2px 8px 0 8px'
          }}>
            <input
              type="text"
              placeholder="Type your prompt..."
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              onKeyPress={(e) => e.key === 'Enter' && handleChatSend()}
              style={{
                flex: 1,
                padding: '8px 12px',
                fontSize: '14px',
                borderRadius: '18px',
                border: '1px solid #d1d1d1',
                marginRight: '6px',
                outline: 'none',
                background: 'linear-gradient(145deg, #e8e8e8, #d4d4d4)',
                transition: 'all 0.3s ease',
                boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
              }}
              onFocus={(e) => {
                e.target.style.borderColor = '#0078d4';
                e.target.style.background = 'linear-gradient(145deg, #f0f0f0, #e0e0e0)';
                e.target.style.boxShadow = '0 4px 12px rgba(0,120,212,0.2)';
              }}
              onBlur={(e) => {
                e.target.style.borderColor = '#d1d1d1';
                e.target.style.background = 'linear-gradient(145deg, #e8e8e8, #d4d4d4)';
                e.target.style.boxShadow = '0 2px 8px rgba(0,0,0,0.1)';
              }}
            />
            <Button
              appearance="primary"
              disabled={loading || !chatInput.trim()}
              onClick={handleChatSend}
              style={{
                borderRadius: '18px',
                minWidth: '55px',
                padding: '8px 12px',
                backgroundColor: '#0078d4',
                color: '#ffffff',
                border: 'none',
                opacity: (loading || !chatInput.trim()) ? 0.5 : 1
              }}
            >
              {loading ? '...' : 'Send'}
            </Button>
          </div>
        )}
      </div>
    </FluentProvider>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;

