using System.Threading;
using WebbIE4;
using System.Windows.Forms;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

// Singleton instance of the browser
frmMain browser = null;

// Start the WebbIE browser in a separate thread
var webbieThread = new Thread(() => {
    browser = new frmMain();
    browser.Show();
    System.Windows.Forms.Application.Run();
});
webbieThread.SetApartmentState(ApartmentState.STA);
webbieThread.Start();

// Wait for the browser to be created
while (browser == null) {
    Thread.Sleep(100);
}

app.MapGet("/api/browser/content", () => {
    string content = "";
    browser.Invoke((MethodInvoker)delegate {
        content = browser.txtText.Text;
    });
    return content;
});

app.MapPost("/api/browser/navigate", (NavigateRequest request) => {
    browser.Invoke((MethodInvoker)delegate {
        browser.StartNavigating(request.Url);
    });
    return $"Navigating to {request.Url}...";
});

app.MapPost("/api/browser/back", () => {
    browser.Invoke((MethodInvoker)delegate {
        browser.DoBack();
    });
    return "Navigating back...";
});

app.MapPost("/api/browser/forward", () => {
    browser.Invoke((MethodInvoker)delegate {
        browser.DoForward();
    });
    return "Navigating forward...";
});

app.MapPost("/api/browser/click", (ClickRequest request) => {
    // This is a simplified implementation. A real implementation would need to
    // find the element by its ID and then simulate a click.
    browser.Invoke((MethodInvoker)delegate {
        // We can't directly click the element from here, so we'll need to add
        // a new public method to frmMain to handle this.
        // For now, we'll just navigate to the link's address.
        if (request.ElementId >= 0 && request.ElementId < modGlobals.gNumLinks)
        {
            var link = modGlobals.gLinks[request.ElementId];
            browser.StartNavigating(link.address);
        }
    });
    return $"Clicking element {request.ElementId}...";
});

app.MapPost("/api/browser/type", (TypeRequest request) => {
    // This is a simplified implementation. A real implementation would need to
    // find the element by its ID and then set its value.
    browser.Invoke((MethodInvoker)delegate {
        // We can't directly type into the element from here, so we'll need to add
        // a new public method to frmMain to handle this.
        if (request.ElementId >= 0 && request.ElementId < modGlobals.numTextInputs)
        {
            var textInput = modGlobals.textInput[request.ElementId];
            textInput.setAttribute("value", request.Text);
        }
    });
    return $"Typing '{request.Text}' into element {request.ElementId}...";
});

app.Run();

// Request models
public record NavigateRequest(string Url);
public record ClickRequest(int ElementId);
public record TypeRequest(int ElementId, string Text);
