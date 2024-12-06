// Get active document
var doc = app.activeDocument;

// Get active spread
var spread = doc.layoutWindows[0].activeSpread;

// Get spread pages
var pages = spread.pages;

// Get document bleed
var bleedTop = doc.documentPreferences.documentBleedTopOffset;
var bleedBottom = doc.documentPreferences.documentBleedBottomOffset;
var bleedInside = doc.documentPreferences.documentBleedInsideOrLeftOffset;
var bleedOutside = doc.documentPreferences.documentBleedOutsideOrRightOffset;

// Get total width and height + bleed for the spread
function getTotalSpreadDimensions(pages) {
    var leftPage = pages[0];
    var rightPage = pages[pages.length - 1];
    var pageWidth = rightPage.bounds[3] - leftPage.bounds[1];
    var pageHeight = leftPage.bounds[2] - leftPage.bounds[0];
    var totalWidth = pageWidth + bleedInside + bleedOutside;
    var totalHeight = pageHeight + bleedTop + bleedBottom;
    return { width: totalWidth, height: totalHeight, pageWidth: pageWidth, pageHeight: pageHeight };
}

// Get selected object
var selectedObject = app.selection[0];

// Check if selection and if it's a valid object
if (selectedObject != null && (selectedObject instanceof Rectangle || selectedObject instanceof Polygon || selectedObject instanceof Oval || selectedObject instanceof Graphic || selectedObject instanceof TextFrame)) {
    var dimensions = getTotalSpreadDimensions(pages);
    
    // Calculate geometric bounds for the selected object to fill spread and the bleed
    var bounds = [-bleedTop, 
                  -bleedInside, 
                  dimensions.pageHeight + bleedBottom, 
                  dimensions.pageWidth + bleedOutside];

    // Transform selected object to fill the spread and the bleed
    selectedObject.geometricBounds = bounds;

    // If the selected object is a graphic, fit the content proportionally
    if (selectedObject instanceof Rectangle && selectedObject.images.length > 0) {
        selectedObject.fit(FitOptions.FILL_PROPORTIONALLY);
    }
} else {
    alert("Select the object that needs to be transformed.");
}