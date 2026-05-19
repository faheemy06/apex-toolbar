# El-Toolbar — Excel Finance Add-in

A minimalist command panel for Microsoft Excel, designed to help 
accounting and finance professionals work more efficiently without 
the clutter of Excel's default ribbon.

## What it does

| Button | Function |
|--------|----------|
| £ | Formats selected cells as £ currency with bold styling |
| ∑ | Smart AutoSum — automatically detects horizontal or vertical range |
| ⋈ | Audit X-Ray — highlights formulas in blue, hardcoded values in orange |
| VAT | Calculates 20% VAT and places the result in the adjacent cell |
| Δ% | Calculates percentage variance between two values |
| ❄ | Freezes the top header row |
| ⌫ | Clears all formatting and fills from selected cells |
| □ | Unfreezes panes |

## Who it is for

Built for two types of Excel user:

- **Beginners** — new to Excel and overwhelmed by the ribbon, who need 
quick access to everyday finance tasks
- **Power users** — experienced professionals who want fewer clicks and 
faster workflows

## How to install

1. Download `manifest.xml` from this repository
2. Open Microsoft Excel — desktop app or Excel Online
3. Click **Insert** → **Add-ins** → **Upload My Add-in**
4. Select the `manifest.xml` file
5. El-Toolbar will appear as a panel in your spreadsheet

## Built with

- HTML5 and CSS3
- JavaScript (ES6) with asynchronous Office.js API calls
- Microsoft Office.js API v1.1
- Hosted on GitHub Pages

## Live URL

https://faheemy06.github.io/El-Toolbar/index.html
