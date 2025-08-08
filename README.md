# WebbIE Modernized Architecture

This document outlines the new software architecture for WebbIE, which has been refactored to use modern technologies, including the `WebView2` browser control and an in-browser Large Language Model (LLM) for assistive navigation.

## Overview

The original WebbIE application was built using the legacy `System.Windows.Forms.WebBrowser` control, which is based on an outdated version of Internet Explorer. This architecture presented significant limitations for implementing modern web features.

The new architecture replaces this core component with **Microsoft's `WebView2` control**, which is based on the Chromium engine. This fundamental change enables modern web standards and allows for deep, robust integration with in-page JavaScript.

The primary goal of this refactoring was to embed an LLM directly into the browser, allowing it to act as an intelligent assistant that can understand user commands and interact with web pages on their behalf.

## Core Components

The application is now structured around the following key components:

-   **`frmMain.vb`**: The primary application window. It hosts the text-based representation of the web page and the main user controls (toolbar, menus). It is responsible for orchestrating the user's interaction with the AI.

-   **`frmWeb.vb`**: A secondary form that hosts the `WebView2` browser control. This form is not always visible to the user but is responsible for all live web rendering and for injecting the necessary JavaScript into the loaded pages.

-   **`WebView2` Control**: The modern browser engine that renders web pages. It provides the necessary APIs for the .NET host to communicate with the JavaScript running on the page.

## The AI / JavaScript Layer

A significant portion of the application's logic has been moved into a dedicated JavaScript layer that runs inside the `WebView2` control.

-   **`dom-parser.js`**: This script is injected into every webpage. Its job is to traverse the Document Object Model (DOM) and extract a clean, flat list of all text and interactable elements (links, buttons, inputs, etc.). It assigns a unique ID to each element and returns this structured data to the .NET host as a JSON object. This provides `frmMain` with the content for its text view.

-   **`Transformers.js`**: This powerful machine learning library is loaded from a CDN. It enables us to download and run sophisticated AI models directly within the user's browser, without requiring a connection to an external server.

-   **`llm-handler.js`**: This is the "brain" of the AI assistant. It uses `Transformers.js` to load a small, quantized language model (e.g., `Xenova/LaMini-Flan-T5-77M`). When the user makes a request, this script constructs a detailed prompt containing the page's content and the user's query, and asks the LLM to decide on the single best action to take.

## The .NET-JavaScript Communication Bridge

To allow the JavaScript code to securely and reliably call back to the .NET host, a formal API bridge has been established using a "Host Object".

-   **`ToolHost.vb`**: This .NET class defines the set of tools that the AI can use. It contains simple, public methods like `click(elementId)`, `type(elementId, text)`, and `goTo(url)`.

-   **`AddHostObjectToScript`**: In `frmWeb.vb`, an instance of the `ToolHost` class is exposed to the JavaScript environment under the name `window.chrome.webview.hostObjects.webbieTools`. This means that when the JavaScript code calls `window.chrome.webview.hostObjects.webbieTools.click(...)`, it is directly invoking the `Click` method on the `ToolHost` object in the .NET application.

## "Ask AI" Workflow

Here is a step-by-step summary of how a user command is processed:

1.  The user clicks the "Ask AI" button in the `frmMain` toolbar.
2.  An input box prompts the user for a command (e.g., "click the login button" or "type 'hello world' in the search bar").
3.  `frmMain` calls the global `window.aiProcess` JavaScript function, passing it the full text content of the page and the user's command.
4.  Inside `llm-handler.js`, the `aiProcess` function formats a prompt for the LLM, asking it to generate the specific `webbieTools` JavaScript call that will achieve the user's goal.
5.  The LLM generates a line of code, for example: `window.chrome.webview.hostObjects.webbieTools.click("webbie-id-42")`.
6.  The `llm-handler.js` script executes this line of code using `eval()`.
7.  The `WebView2` control marshals this call across the boundary from JavaScript to .NET, invoking the `Click` method on the `ToolHost` object.
8.  The `ToolHost.Click` method then calls a public method on `frmMain` (`ExecuteClick`), which runs the final JavaScript snippet to perform the click on the webpage and refresh the text view.

## How to Extend the AI's Capabilities

Adding a new tool for the AI to use is straightforward:

1.  **Add a new public method** to the `ToolHost.vb` class (e.g., `Public Sub Scroll(direction As String)`).
2.  **Add the corresponding execution logic** as a public method in `frmMain.vb` (e.g., `Public Sub ExecuteScroll(direction As String)`).
3.  **Update the prompt** in `llm-handler.js` to include the new tool in the list of options available to the LLM (e.g., `window.chrome.webview.hostObjects.webbieTools.scroll("down")`).
4.  **Optionally**, add a new `Case` in the `btnAskAI_Click` handler in `frmMain.vb` if the tool requires complex argument handling (though for most tools, this is not necessary).
