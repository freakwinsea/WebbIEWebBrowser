// llm-handler.js

// Define a class to encapsulate the AI functionality
class AIHandler {
    static instance = null;

    constructor() {
        // The model pipeline will be initialized once
        this.pipeline = null;
    }

    // Singleton pattern to ensure we only initialize the pipeline once
    static async getInstance() {
        if (this.instance === null) {
            this.instance = new AIHandler();
            // Important: The library is loaded from the CDN in frmWeb.vb
            // We need to make sure it's loaded before we try to use it.
            // The 'transformers' object will be available globally.
            const { pipeline } = await transformers;
            this.instance.pipeline = await pipeline('text2text-generation', 'Xenova/LaMini-Flan-T5-77M', {
                quantized: true, // Use a quantized model for speed and memory efficiency
            });
        }
        return this.instance;
    }

    // The main function to be called from the C# code
    async process(pageContent, userQuery) {
        if (!this.pipeline) {
            console.error("Pipeline not initialized!");
            return JSON.stringify({ error: "Pipeline not initialized" });
        }

        // Construct the prompt for the LLM.
        // This prompt now asks for a direct JavaScript call to the exposed .NET object.
        const prompt = `
You are an expert web navigation assistant for users with low or no vision.
Your task is to analyze the provided webpage content and the user's request, then decide which single action to take.
The webpage content lists all interactable elements with a unique ID, like "LINK [webbie-id-12]: Sign in".

You must choose one of the following tools by generating the full JavaScript code to call it:
1. window.chrome.webview.hostObjects.webbieTools.click("element_id"): Use this to click on a link, button, or other element. Use the full element_id string (e.g., "webbie-id-12").
2. window.chrome.webview.hostObjects.webbieTools.type("element_id", "text to type"): Use this to type text into an input field.
3. window.chrome.webview.hostObjects.webbieTools.goTo("url"): Use this to navigate to a new URL.
4. window.chrome.webview.hostObjects.webbieTools.answer("text"): Use this to provide a direct answer to the user's question if no navigation or interaction is needed.

Based on the following webpage content and user query, what is the single best tool call to execute?
You MUST respond with ONLY the single line of JavaScript code and nothing else.

--- WEBPAGE CONTENT ---
${pageContent}
-----------------------

--- USER QUERY ---
${userQuery}
------------------

JAVASCRIPT:`;

        try {
            // Generate the response from the model
            const response = await this.pipeline(prompt, {
                max_new_tokens: 100, // Increased tokens to allow for longer URLs or text
                do_sample: false, // Be more deterministic
            });

            const generatedCode = response[0].generated_text.trim();

            // For security and robustness, we should validate the generated code
            // to ensure it's a call to our exposed API.
            if (generatedCode.startsWith("window.chrome.webview.hostObjects.webbieTools.")) {
                console.log("Executing AI code:", generatedCode);
                // Execute the generated code
                eval(generatedCode);
                // Return a success message or void, as the action is now handled by the host.
                return JSON.stringify({ success: true, executed: generatedCode });
            } else {
                // If the model generates something else, treat it as a direct answer.
                console.log("AI returned a direct answer:", generatedCode);
                window.chrome.webview.hostObjects.webbieTools.answer(generatedCode);
                return JSON.stringify({ success: true, executed: `answer("${generatedCode}")` });
            }

        } catch (error) {
            console.error("Error during pipeline execution:", error);
            return JSON.stringify({ success: false, error: error.message });
        }
    }
}

// Make the process function globally accessible so C# can call it.
// We'll attach it to the window object.
window.aiProcess = async (pageContent, userQuery) => {
    const handler = await AIHandler.getInstance();
    return await handler.process(pageContent, userQuery);
};
