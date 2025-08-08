# Connecting an LLM to WebbIE

This document explains how to connect a Large Language Model (LLM) to the WebbIE browser.

## Overview

The integration works by running a local web server that exposes an API for controlling the WebbIE browser. Your LLM can make requests to this API to perform actions like navigating to a URL, getting the page content, and interacting with elements on the page.

## Prerequisites

*   You have an LLM that can make HTTP requests.
*   You have built and are running the `WebbIE.Api` project.

## Running the API Server

1.  Open a terminal or command prompt.
2.  Navigate to the `WebbIE.Api` directory.
3.  Run the command `dotnet run`.

The API server will start and listen on `http://localhost:5000`.

## API Endpoints

The following API endpoints are available:

### Get Page Content

*   **Endpoint:** `GET /api/browser/content`
*   **Description:** Returns the parsed text content of the current page.
*   **Response:** A plain text string representing the page content.

### Navigate to a URL

*   **Endpoint:** `POST /api/browser/navigate`
*   **Description:** Navigates the browser to the specified URL.
*   **Request Body:** A JSON object with a `url` property.
    ```json
    {
        "url": "https://www.example.com"
    }
    ```

### Go Back

*   **Endpoint:** `POST /api/browser/back`
*   **Description:** Navigates to the previous page in the browser's history.

### Go Forward

*   **Endpoint:** `POST /api/browser/forward`
*   **Description:** Navigates to the next page in the browser's history.

### Click an Element

*   **Endpoint:** `POST /api/browser/click`
*   **Description:** Clicks on a specific interactable element on the page.
*   **Request Body:** A JSON object with an `elementId` property. The `elementId` corresponds to the index of the element in the list of interactable elements.
    ```json
    {
        "elementId": 0
    }
    ```

### Type Text

*   **Endpoint:** `POST /api/browser/type`
*   **Description:** Types text into a specific input field on the page.
*   **Request Body:** A JSON object with `elementId` and `text` properties.
    ```json
    {
        "elementId": 0,
        "text": "Hello, world!"
    }
    ```

## Example Usage with an LLM

Here is a conceptual example of how an LLM could use this API to search for "WebbIE" on Google:

1.  **LLM:** "I need to search for 'WebbIE' on Google."
2.  **LLM sends request:** `POST /api/browser/navigate` with `{"url": "https://www.google.com"}`.
3.  **WebbIE navigates to Google.**
4.  **LLM sends request:** `GET /api/browser/content` to get the page content.
5.  **LLM finds the search input field** in the page content and determines its `elementId`.
6.  **LLM sends request:** `POST /api/browser/type` with `{"elementId": <search_field_id>, "text": "WebbIE"}`.
7.  **WebbIE types "WebbIE" into the search field.**
8.  **LLM finds the "Google Search" button** and determines its `elementId`.
9.  **LLM sends request:** `POST /api/browser/click` with `{"elementId": <search_button_id>}`.
10. **WebbIE clicks the search button.**
11. **LLM sends request:** `GET /api/browser/content` to get the search results.

This example demonstrates the basic workflow of using the API to automate web browsing tasks.
