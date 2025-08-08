import os
import subprocess
import time
from playwright.sync_api import sync_playwright, expect

def run_e2e_test():
    with sync_playwright() as p:
        # Define the path to the WebbIE executable
        # This assumes the script is run from the repository root
        executable_path = os.path.join(os.getcwd(), "WebbIE4.NET", "bin", "Debug", "WebbIE4.NET.exe")

        if not os.path.exists(executable_path):
            print(f"Error: Executable not found at {executable_path}")
            print("Please ensure the WebbIE4.NET project has been built in Debug mode.")
            return

        # Launch the WebbIE application as a subprocess
        print(f"Launching {executable_path}...")
        proc = subprocess.Popen([executable_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        try:
            # Wait for the WebView2 to signal that it's initialized
            print("Waiting for WebView2 to initialize...")
            initialized = False
            for line in iter(proc.stdout.readline, ''):
                print(f"WebbIE stdout: {line.strip()}")
                if "WebView2 initialized" in line:
                    initialized = True
                    print("WebView2 initialized successfully.")
                    break
                # Add a timeout condition
                if "Error" in line:
                    print(f"Error during WebbIE startup: {line.strip()}")
                    break

            if not initialized:
                print("Error: Timed out waiting for WebView2 initialization signal.")
                return

            # Connect to the running WebView2 instance via CDP
            print("Connecting to WebView2 via CDP...")
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
            page = context.pages[0]
            print("Successfully connected to page.")

            # Navigate to the local test file
            test_page_path = os.path.join(os.getcwd(), "WebbIE4.NET", "Tests", "mock_pages", "simple_form.html")
            test_page_uri = f"file:///{test_page_path.replace(os.sep, '/')}"
            print(f"Navigating to {test_page_uri}...")
            page.goto(test_page_uri)

            # Give the page a moment to load and our scripts to run
            page.wait_for_load_state('domcontentloaded')
            time.sleep(2) # Allow extra time for scripts to be ready

            # Use the AI to type into the username field
            print("Executing AI command...")
            command = "type 'jules-tester' into the username field"
            # Note: We don't need the DOM content for this test since the AI is simple
            # and the command is direct. In a real scenario, we'd pass the parsed DOM.
            page.evaluate(f"window.aiProcess('', `{command}`)")

            # Wait for the action to be processed
            time.sleep(1)

            # Assert that the username field was filled correctly
            print("Asserting final state...")
            username_input = page.locator("#username")
            expect(username_input).to_have_value("jules-tester")
            print("Assertion passed!")

            # Take a screenshot for visual verification
            screenshot_path = "jules-scratch/verification/e2e_test_result.png"
            print(f"Taking screenshot: {screenshot_path}")
            page.screenshot(path=screenshot_path)

        finally:
            # Clean up: close the browser connection and terminate the process
            print("Cleaning up...")
            if 'browser' in locals() and browser.is_connected():
                browser.close()
            proc.terminate()
            proc.wait()
            print("Test run finished.")

if __name__ == "__main__":
    run_e2e_test()
