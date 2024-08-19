// Define variables
var doc = app.activeDocument;
var startPage = 16;
var endPage = 25;
var registration = doc.swatches.item("Registration");
var newBlack = doc.swatches.item("Black");
var whiteSwatch = doc.swatches.item("White");

// Main function to replace colors in the specified page range
function replaceColors(items, registration, newBlack, whiteSwatch) {
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    try {
      if (item.fillColor && item.fillColor.name === registration.name && (whiteSwatch === null || item.fillColor.name !== whiteSwatch.name)) {
        item.fillColor = newBlack;
      }
      if (item.strokeColor && item.strokeColor.name === registration.name && (whiteSwatch === null || item.strokeColor.name !== whiteSwatch.name)) {
        item.strokeColor = newBlack;
      }
      if (item.allPageItems && item.allPageItems.length > 0) {
        replaceColors(item.allPageItems, registration, newBlack, whiteSwatch);
      }
    } catch (e) {
      // Log error for debugging
      $.writeln("Error processing item: " + e.message);
    }
  }
}

function replaceColorsInPageRange(doc, startPage, endPage, registration, newBlack, whiteSwatch) {
  for (var i = startPage - 1; i < endPage; i++) {
    var page = doc.pages[i];
    replaceColors(page.allPageItems, registration, newBlack, whiteSwatch);
    replaceTextColors(page.textFrames, registration, newBlack, whiteSwatch);
  }
  replaceCharacterStyles(doc.allCharacterStyles, registration, newBlack, whiteSwatch);
  replaceParagraphStyles(doc.allParagraphStyles, registration, newBlack, whiteSwatch);
}

function replaceTextColors(textFrames, registration, newBlack, whiteSwatch) {
  for (var i = 0; i < textFrames.length; i++) {
    var textFrame = textFrames[i];
    if (textFrame.fillColor && textFrame.fillColor.name === registration.name && (whiteSwatch === null || textFrame.fillColor.name !== whiteSwatch.name)) {
      if (!(textFrame.parentStory.contents === "KDP White")) {
        textFrame.fillColor = newBlack;
      }
    }
    if (textFrame.strokeColor && textFrame.strokeColor.name === registration.name && (whiteSwatch === null || textFrame.strokeColor.name !== whiteSwatch.name)) {
      if (!(textFrame.parentStory.contents === "KDP White")) {
        textFrame.strokeColor = newBlack;
      }
    }
  }
}

function replaceCharacterStyles(characterStyles, registration, newBlack, whiteSwatch) {
  for (var i = 0; i < characterStyles.length; i++) {
    var style = characterStyles[i];
    if (style.fillColor && style.fillColor.name === registration.name && (whiteSwatch === null || style.fillColor.name !== whiteSwatch.name)) {
      style.fillColor = newBlack;
    }
    if (style.strokeColor && style.strokeColor.name === registration.name && (whiteSwatch === null || style.strokeColor.name !== whiteSwatch.name)) {
      style.strokeColor = newBlack;
    }
  }
}

function replaceParagraphStyles(paragraphStyles, registration, newBlack, whiteSwatch) {
  for (var i = 0; i < paragraphStyles.length; i++) {
    var style = paragraphStyles[i];
    if (style.fillColor && style.fillColor.name === registration.name && (whiteSwatch === null || style.fillColor.name !== whiteSwatch.name)) {
      style.fillColor = newBlack;
    }
    if (style.strokeColor && style.strokeColor.name === registration.name && (whiteSwatch === null || style.strokeColor.name !== whiteSwatch.name)) {
      style.strokeColor = newBlack;
    }
  }
}

// Call the main function with the established variables
replaceColorsInPageRange(doc, startPage, endPage, registration, newBlack, whiteSwatch);