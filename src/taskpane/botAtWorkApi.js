// Gemini API utility for suggesting replies
const GEMINI_API_KEY = "AIzaSyAC6XVMIRh5CzUjqKPu8Y_A19iPCZNfTdc";

const urlProd = "https://mail-magicplugin-p872js3nd-subhasisnayak270-5522s-projects.vercel.app/";

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
    'timeout'
  ];
  
  const errorMessage = (error.message || '').toLowerCase();
  return retryableMessages.some(msg => errorMessage.includes(msg));
};

export async function getSuggestedReply(emailBody, maxRetries = 3) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + GEMINI_API_KEY;
  const body = {
    contents: [
      {
        parts: [
          {
            text: emailBody,
          },
        ],
      },
    ],
  };

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`Gemini API attempt ${attempt}/${maxRetries}`);
      
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      
      const data = await response.json();
      console.log("Gemini API raw response:", data);
      
      // Check for successful response
      if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) {
        return data.candidates[0].content.parts.map((p) => p.text).join("\n");
      }
      
      // Check for API errors
      if (data.error) {
        const errorMessage = data.error.message || JSON.stringify(data.error);
        console.log(`Gemini API error on attempt ${attempt}:`, errorMessage);
        
        // Check if this is a retryable error
        if (isRetryableError(data.error) && attempt < maxRetries) {
          const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000; // Exponential backoff with jitter
          console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
          await delay(backoffDelay);
          continue;
        }
        
        // If not retryable or max retries reached, return error
        return `Gemini API error: ${errorMessage}${attempt > 1 ? ` (after ${attempt} attempts)` : ''}`;
      }
      
      // No candidates returned
      return "No suggestion returned. Raw response: " + JSON.stringify(data);
      
    } catch (e) {
      console.log(`Network error on attempt ${attempt}:`, e.message);
      
      // Check if this is a retryable network error
      if (attempt < maxRetries) {
        const backoffDelay = Math.pow(2, attempt - 1) * 1000 + Math.random() * 1000;
        console.log(`Retrying in ${Math.round(backoffDelay)}ms...`);
        await delay(backoffDelay);
        continue;
      }
      
      // Max retries reached
      return `Error calling Gemini API: ${e.message} (after ${attempt} attempts)`;
    }
  }
  
  // This should never be reached, but just in case
  return "Maximum retry attempts reached. Please try again later.";
}
