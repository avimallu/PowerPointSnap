# Snap for Microsoft PowerPoint
A simple VBA based add-in to automate a few things that are missing in PowerPoint. Can apply properties (size, colour, axes) of Object A to another Object B. If there are bugs, feel free to raise an issue. I'm not a professional programmer, so I'll look into it when I get the time.

## Features

### Decent, one click UI

![SnapUI](https://raw.githubusercontent.com/avimallu/PowerPointSnap/master/Screenshots/SnapScreenshot.JPG)

### For most objects
* *Snap* an object's properties to another one.
* Support typical shapes (Insert > Add Shapes), Text Boxes (both normal and default ones in a layout), Charts, and Tables.
* Supports copying the size (height and width), colour (fill and outline), position of most shapes and text boxes.

### For charts
* Supports syncing the axes of charts with differing axis lengths, plot areas, title positions and legend positions (both charts must support the property).
* Supports data label customization options - highlight only specific (or all) legend entries, for their first/last/min/max/median or every N values.

![DataLabelsUI](https://raw.githubusercontent.com/avimallu/PowerPointSnap/master/Screenshots/DataLabelsScreenshot.JPG)

### For tables
* Supports syncing row heights and column widths to tables with identical row/column counts (even with complicated merging in place).

### For enterprise users
* Available as an VBA Add-in, requiring no installation and no administrative rights on most enterprise systems.

## Installation
* Download the `Snap.ppam` file to your system.
* [Enable](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45) the developer options.
* Go to the **Developer** tab, and click on **PowerPoint Add-ins**.
* Click on *Add New*. Choose the location of the file you just dowloaded. Click *Close*.
* To **uninstall**, repeat the process, and simply click on *Remove* this time.

## Privacy?
**This code by itself does not collet any data whatsoever**. Whatever you're working on is for your eyes only, at least from the tool's perspective.

## License
In a gist, I don't care what you do with it, as long as you don't expect it to work all the time and don't blame me for any consequence arising from its use. Offically, read the license file in the repository.
