// Get active document
var doc = app.activeDocument;

// Get active spread
var spread = doc.layoutWindows[0].activeSpread;

// Get spread pages
var pages = spread.pages;

// Get document margins
function getPageMargins(page) {
    return {
        top: page.marginPreferences.top,
        bottom: page.marginPreferences.bottom,
        inside: page.marginPreferences.left,
        outside: page.marginPreferences.right
    };
}

// Get total width and height limited to margins for the spread
function getTotalSpreadDimensions(pages) {
    var leftPage = pages[0];
    var rightPage = pages[pages.length - 1];
    var pageWidth = rightPage.bounds[3] - leftPage.bounds[1];
    var pageHeight = leftPage.bounds[2] - leftPage.bounds[0];

    var leftPageMargins = getPageMargins(leftPage);
    var rightPageMargins = getPageMargins(rightPage);

    var totalWidth = pageWidth - (leftPageMargins.inside + rightPageMargins.outside);
    var totalHeight = pageHeight - Math.max(leftPageMargins.top, rightPageMargins.top) 
                                    - Math.max(leftPageMargins.bottom, rightPageMargins.bottom);

    return { 
        width: totalWidth, 
        height: totalHeight, 
        pageWidth: pageWidth, 
        pageHeight: pageHeight,
        leftPageMargins: leftPageMargins,
        rightPageMargins: rightPageMargins
    };
}

// Get selected object
var selectedObject = app.selection[0];

// Check if selection and if it's a valid object
if (selectedObject != null && (selectedObject instanceof Rectangle || selectedObject instanceof Polygon || selectedObject instanceof Oval || selectedObject instanceof Graphic || selectedObject instanceof TextFrame)) {
    var dimensions = getTotalSpreadDimensions(pages);
    var leftPage = pages[0];
    
    // Calculate geometric bounds for the selected object to fill the page area within margins
    var bounds = [
        dimensions.leftPageMargins.top, 
        dimensions.leftPageMargins.inside, 
        dimensions.pageHeight - dimensions.leftPageMargins.bottom, 
        dimensions.pageWidth - dimensions.rightPageMargins.outside
    ];

    // Transform selected object to fill the spread area within margins
    selectedObject.geometricBounds = bounds;

    // If the selected object is a graphic, fit the content proportionally
    if (selectedObject instanceof Rectangle && selectedObject.images.length > 0) {
        selectedObject.fit(FitOptions.FILL_PROPORTIONALLY);
    }
} else {
    alert("Select the object that needs to be transformed.");
}
