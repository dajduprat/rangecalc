# RangeCalc - Excel Custom Functions Add-in

A modern Excel add-in that provides custom functions for range calculations and data manipulation. Built with Office JavaScript API and bundled with Webpack.

## Features

- **Custom Functions**: Write Excel-compatible functions in JavaScript
- **Range Calculation**: Calculate and manipulate Excel ranges
- **Safe Calculation Mode**: Error-handling enabled calculation with automatic recovery
- **Entity Support**: Work with Excel EntityCellValue for structured data
- **Task Pane UI**: Interactive task pane for add-in controls
- **Ribbon Commands**: Quick access buttons in the Excel ribbon

## Project Structure

```
├── src/
│   ├── commands/          # Ribbon command handlers
│   ├── functions/         # Custom function definitions
│   ├── taskpane/          # Task pane UI (HTML, CSS, JS)
│   └── utils/             # Utility functions
├── assets/                # Icon assets (PNG format)
├── manifest.xml           # Add-in configuration
├── webpack.config.js      # Build configuration
├── babel.config.json      # JavaScript transpilation
└── package.json           # Dependencies and scripts
```

## Prerequisites

- **Node.js**: Version 16+ (LTS recommended). [Download](https://nodejs.org/)
- **Office**: Microsoft 365 subscription with Excel installed
- **Modern Browser**: Edge Chromium or later for debugging

## Quick Start

### 1. Install Dependencies

```bash
npm install
```

### 2. Start Development Server

```bash
npm run dev-server
```

The server runs on `https://localhost:3000/`

### 3. Debug the Add-in

Use the Office Add-ins Development Kit extension in VS Code:
- Press **F5** or select **Preview Your Office Add-in (F5)**
- Choose **Excel Desktop (Edge Chromium)**
- The add-in will sideload into Excel automatically

### 4. Build for Production

```bash
npm run build
```

This creates optimized files in the `dist/` directory with all localhost URLs replaced with your production domain (`TBD`).

## Available Scripts

| Command | Purpose |
|---------|---------|
| `npm run build:dev` | Development build with source maps |
| `npm run build` | Production build (minified, optimized) |
| `npm run dev-server` | Start HTTPS dev server |
| `npm run watch` | Watch files and rebuild on changes |
| `npm run start` | Debug add-in in Excel (F5 equivalent) |
| `npm run stop` | Stop debugging session |
| `npm run lint` | Check code for issues |
| `npm run lint:fix` | Auto-fix linting issues |

## Developing Custom Functions

Add functions to `src/functions/functions.js`:

```javascript
/**
 * Adds two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} Sum of the numbers
 */
export function add(first, second) {
  return first + second;
}
```

The `@customfunction` decorator exposes the function to Excel. Rebuild to see changes in Excel.

## Ribbon Commands

Ribbon buttons are defined in `manifest.xml` and handled in `src/commands/commands.js`:

- **Calculate Range** (Red ✗): Fast calculation without error recovery
- **Safe Calculate** (Green ✓): Calculation with error handling and stale reference recovery

## Entity Cell Values

The add-in supports Excel's `EntityCellValue` type for structured data:

```javascript
const entity = {
  type: "Entity",
  text: "Display Text",
  properties: {
    propertyName: { type: "String", basicValue: "value" },
    propertyId: { type: "Double", basicValue: 123 },
    isActive: { type: "Boolean", basicValue: true }
  },
  layouts: {
    card: {
      title: { property: "propertyName" },
      sections: [...]
    }
  }
};
```

## Troubleshooting

### Add-in not loading
- Close all Excel instances
- Run `npm run stop` to clean up previous sessions
- Check the dev server is running: `npm run dev-server`

### Manifest validation errors
- Use **Validate Manifest File** in Office Add-ins Development Kit
- Ensure all icon paths are correct in `manifest.xml`

### Custom functions not appearing
- Rebuild with `npm run build:dev`
- Reload Excel and the add-in
- Check `functions.json` is generated in `dist/`

### Build not replacing localhost URLs
- Verify `urlProd` in `webpack.config.js` is set correctly
- Run `npm run build` (not dev build)
- Check output files in `dist/` directory

## Configuration

### Update Production URL

Edit `webpack.config.js`:

```javascript
const urlProd = "https://your-domain.com/"; // Change this
```

### Manifest Settings

Key settings in `manifest.xml`:
- `<IconUrl>`: Task pane icon
- `<HighResolutionIconUrl>`: High DPI icon
- `<SupportUrl>`: Support link
- `<SourceLocation>`: Task pane HTML location

## Testing

### Local Testing
- Use `npm run dev-server` + Excel debug mode
- Monitor console output in Edge DevTools
- Check `functions.json` generation for custom functions

### Production Testing
- Build with `npm run build`
- Deploy files to your production server
- Test with actual domain URLs

## Resources

- [Excel Custom Functions Documentation](https://learn.microsoft.com/office/dev/add-ins/excel/custom-functions-overview)
- [Office JavaScript API Reference](https://learn.microsoft.com/javascript/api/overview)
- [EntityCellValue Type](https://learn.microsoft.com/javascript/api/excel/excel.entitycellvalue)
- [Office Add-ins Dev Kit](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-development-kit)

## License

TBD

## Disclaimer

THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
