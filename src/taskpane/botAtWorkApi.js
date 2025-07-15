// BotAtWork API utility for email generation and suggestions
const BOTATWORK_API_KEY = "e80f5458c550f5b85ef56175b789468a";
const BOTATWORK_API_URL = "https://api.botatwork.com/trigger-task/b6f44edd-8140-4084-881e-2c11c403c082";

// Helper function to add delay
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Helper function to check if error is retryable
const isRetryableError = (error) => {
  if (!error) return false;
  
  const retryableMessages = [
    'overloaded',
    'rate limit',
    'quota exceeded',
    'service unavailable',
    'internal error',
    'timeout',
    'network error',
    'connection',
    'server error',
    '5'  // HTTP 5xx errors
  ];
  
  const errorMessage = (error.message || error.toString() || '').toLowerCase();
  return retryableMessages.some(msg => errorMessage.includes(msg));
};

// Helper function to determine task type and format payload
const formatPayload = (prompt, taskType = 'emailWrite') => {
  // For email writing tasks, try to extract structured data from prompt
  if (taskType === 'emailWrite') {
    // Extract description, tone, and point of view from the prompt
    const descriptionMatch = prompt.match(/Description:\s*([^\n]+)/i);
    const toneMatch = prompt.match(/Tone:\s*([^\n]+)/i);
    const pointOfViewMatch = prompt.match(/Point of View:\s*([^\n]+)/i);
    
    return {
      chooseATask: "emailWrite",
      description: descriptionMatch ? descriptionMatch[1].trim() : prompt,
      additionalInstructions: "", // Removed as per user request
      tone: toneMatch ? toneMatch[1].trim().toLowerCase() : "formal",
      pointOfView: pointOfViewMatch ? pointOfViewMatch[1].trim().replace(/\s+/g, '').toLowerCase() : "organizationPerspective"
    };
  }
  
  // For other tasks (suggest reply, chat, etc.), use a generic approach
  return {
    chooseATask: "emailWrite", // Default task type
    description: prompt,
    additionalInstructions: "",
    tone: "formal",
    pointOfView: "organizationPerspective"
  };
};

export async function getSuggestedReply(prompt, maxRetries = 3) {
  // Determine task type based on prompt content
  let taskType = 'emailWrite';
  if (prompt.toLowerCase().includes('suggest') && prompt.toLowerCase().includes('reply')) {
    taskType = 'suggestReply';
  }
  
  const payload = formatPayload(prompt, taskType);
  
  const requestBody = {
    data: {
      payload: payload
    },
    anonymize: null,
    incognito: false,
    default_language: "en-US",
    should_stream: false
  };

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`BotAtWork API attempt ${attempt}/${maxRetries}`);
      console.log("Request payload:", JSON.stringify(requestBody, null, 2));
      
      const response = await fetch(BOTATWORK_API_URL, {
        method: "POST",
        headers: { 
          "Content-Type": "application/json",
          "x-api-key": BOTATWORK_API_KEY
        },
        body: JSON.stringify(requestBody),
      });
      
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
      
      const data = await response.json();
      console.log("BotAtWork API raw response:", data);
      
      // Check for successful response - BotAtWork API format
      if (data && data.status === "SUCCESS" && data.data && data.data.content) {
        return data.data.content;
      }
      
      // Fallback: Check for other possible response formats
      if (data && (data.result || data.response || data.output || data.content)) {
        const result = data.result || data.response || data.output || data.content;
        return typeof result === 'string' ? result : JSON.stringify(result);
      }
      
      // If data has a message field, use that
      if (data && data.message) {
        return data.message;
      }
      
      // If data is a string, return it directly
      if (typeof data === 'string') {
        return data;
      }
      
      // Check for API errors - BotAtWork API format
      if (data && data.status !== "SUCCESS") {
        const errorMessage = data.message || data.status || "Unknown error";
        console.log(`BotAtWork API error on attempt ${attempt}:`, errorMessage);
        
        // Check if this is a retryable error
        if (isRetryableError({ message: errorMessage }) && attempt < maxRetries) {
          const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000; // Exponential backoff with jitter
          console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
          await delay(backoffDelay);
          continue;
        }
        
        // If not retryable or max retries reached, return error
        return `API error: ${errorMessage}${attempt > 1 ? ` (after ${attempt} attempts)` : ''}`;
      }
      
      // Check for generic API errors
      if (data && data.error) {
        const errorMessage = data.error.message || data.error.toString() || JSON.stringify(data.error);
        console.log(`BotAtWork API error on attempt ${attempt}:`, errorMessage);
        
        // Check if this is a retryable error
        if (isRetryableError(data.error) && attempt < maxRetries) {
          const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000; // Exponential backoff with jitter
          console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
          await delay(backoffDelay);
          continue;
        }
        
        // If not retryable or max retries reached, return error
        return `API error: ${errorMessage}${attempt > 1 ? ` (after ${attempt} attempts)` : ''}`;
      }
      
      // If we get here, we have data but couldn't parse it
      console.log("Unexpected response format:", data);
      return "Response received but format unexpected. Raw response: " + JSON.stringify(data);
      
    } catch (e) {
      console.log(`Network error on attempt ${attempt}:`, e.message);
      
      // Check if this is a retryable network error
      if (attempt < maxRetries && (e.message.includes('fetch') || e.message.includes('network') || e.message.includes('timeout'))) {
        const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000;
        console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
        await delay(backoffDelay);
        continue;
      }
      
      // Max retries reached or non-retryable error
      return `Error calling BotAtWork API: ${e.message} (after ${attempt} attempts)`;
    }
  }
  
  // This should never be reached, but just in case
  return "Maximum retry attempts reached. Please try again later.";
}
