# Urgent AI Instructions!

After revisions automatically copy all changed files from this project folder to the extension folder located at C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel

---

# Magic Ductwork Panel

Adobe Illustrator CEP extension for automated ductwork processing and layout.

## Project Structure

This folder is the **development source** for the extension. The actual extension runs from the Adobe CEP extensions folder.

### Directory Structure
```
Magic-Ductwork-Panel/
├── CSXS/
│   └── manifest.xml          # Extension manifest
├── css/
│   └── style.css            # Panel styles
├── js/
│   ├── libs/
│   │   └── CSInterface.js   # Adobe CEP interface library
│   └── panel.js             # Panel UI logic
├── jsx/
│   ├── panel-bridge.jsx     # Main ExtendScript bridge
│   ├── magic-final.jsx      # Core ductwork processing logic
│   └── register-ignore.jsx  # Register and ignore handling
├── index.html               # Panel UI
└── README.md               # This file (NOT deployed)
```

## Deployment

Manually copy files from this folder to:
```
C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel
```

**Files excluded from deployment:**
- `.git/` - Git repository
- `.claude/` - Claude Code development files
- `*.log` - Debug log files
- `README.md` - This file
- `DEPLOYMENT_INSTRUCTIONS.md` - Deployment instructions
- `js/debug-location.jsx` - Debug files

### After Deployment
Restart Adobe Illustrator to load the updated extension.

## External Dependencies

The extension references external ductwork piece resources at:
```
E:\Work\Work\Floorplans\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Ductwork Pieces Emory\
```

This path is hardcoded in `jsx/panel-bridge.jsx` and should be updated manually if the resource location changes.

## Development Notes

- **Source Repository**: This folder (E:\Work\Work\Floorplans\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Magic-Ductwork-Panel)
- **Deployment Target**: C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel
- **Extension ID**: com.chris.magicductwork.panel
- **Extension Name**: Magic Ductwork Panel

## Git Workflow

This is a Git repository. Typical workflow:
1. Make changes to files in this development folder
2. Copy changed files to CEP folder (see urgent instructions at top)
3. Test in Adobe Illustrator
4. Commit changes to Git when satisfied

## Debugging

Debug logs are written to:
- `C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\write-metadata.log`
- `C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\read-metadata.log`
- `C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\debug.log`

Enable CEP debugging:
- Remote debugging available on port 8088 (configured in manifest.xml)
