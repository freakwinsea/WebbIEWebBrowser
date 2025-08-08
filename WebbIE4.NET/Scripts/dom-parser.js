function parseDomForWebbIE() {
    const interactableElements = [];
    let elementCounter = 0;

    function getElementInfo(element, elementType) {
        elementCounter++;
        const elementId = `webbie-id-${elementCounter}`;
        element.setAttribute('data-webbie-id', elementId);

        return {
            id: elementId,
            type: elementType,
            text: element.innerText.trim() || element.value || element.ariaLabel || '',
            href: element.href || ''
        };
    }

    // Find all links
    document.querySelectorAll('a[href]').forEach(el => {
        interactableElements.push(getElementInfo(el, 'link'));
    });

    // Find all buttons
    document.querySelectorAll('button, input[type="button"], input[type="submit"], input[type="reset"]').forEach(el => {
        interactableElements.push(getElementInfo(el, 'button'));
    });

    // Find all text inputs
    document.querySelectorAll('input[type="text"], input[type="search"], input[type="email"], input[type="password"], textarea').forEach(el => {
        interactableElements.push(getElementInfo(el, 'textInput'));
    });

    // Find all select dropdowns
    document.querySelectorAll('select').forEach(el => {
        interactableElements.push(getElementInfo(el, 'select'));
    });

    // Find all checkboxes
    document.querySelectorAll('input[type="checkbox"]').forEach(el => {
        interactableElements.push(getElementInfo(el, 'checkbox'));
    });

    // Find all radio buttons
    document.querySelectorAll('input[type="radio"]').forEach(el => {
        interactableElements.push(getElementInfo(el, 'radio'));
    });

    return JSON.stringify(interactableElements);
}
