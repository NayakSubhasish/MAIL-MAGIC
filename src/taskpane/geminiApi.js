// Gemini API utility for suggesting replies
const GEMINI_API_KEY = "AIzaSyAC6XVMIRh5CzUjqKPu8Y_A19iPCZNfTdc";

const urlProd = "https://mail-magicplugin-p872js3nd-subhasisnayak270-5522s-projects.vercel.app/";

export async function getSuggestedReply(emailBody) {
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
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const data = await response.json();
    console.log("Gemini API raw response:", data);
    if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) {
      return data.candidates[0].content.parts.map((p) => p.text).join("\n");
    }
    if (data.error) {
      return "Gemini API error: " + (data.error.message || JSON.stringify(data.error));
    }
    return "No suggestion returned. Raw response: " + JSON.stringify(data);
  } catch (e) {
    return "Error calling Gemini API: " + e.message;
  }
}
