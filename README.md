# Urgent AI Instructions!

Before making or deploying changes, read this file, `DEPLOYMENT_INSTRUCTIONS.md`, and `cpp-plugin/AI Info.md` before claiming deployment is complete. Do not assume another project uses the same deploy path.

After revisions automatically copy all changed files from this project folder to the extension folder located at C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel

If the change touches the Emory/C++ plugin path, `cpp-plugin/`, or anything that affects `ProcessDuctwork.aip`, do not stop after editing or building. You must run the scheduled task `Reload Illustrator Ductwork` so the elevated reload script rebuilds the plugin, copies the `.aip` into `C:\Program Files\Adobe\Adobe Illustrator 2024\Plug-ins\DuctworkMenu\`, syncs the CEP extension, and launches Illustrator for testing.

Important: the `Process Emory Ductwork` button in this panel does not use `ProcessDuctwork.aip`. It calls the separate `EmoryDuctwork` plugin from the `Emory-Ductwork-Panel` project and deploys `C:\Program Files\Adobe\Adobe Illustrator 2024\Plug-ins\DuctworkMenu\EmoryDuctwork.aip`. Do not debug Emory Mode by patching `Process Ductwork` only.

RUN THIS AFTER PLUGIN CHANGES:

```powershell
schtasks /run /tn "Reload Illustrator Ductwork"
```

Wrapper script:

```powershell
powershell -ExecutionPolicy Bypass -File ".\tools\reload-illustrator-ductwork.ps1"
```

Deployment is not complete until all of these are true:
- The task `Reload Illustrator Ductwork` finished with `Last Result: 0`
- `C:\Program Files\Adobe\Adobe Illustrator 2024\Plug-ins\DuctworkMenu\ProcessDuctwork.aip` was updated
- The installed plugin hash matches `E:\Work\Work\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Process Ductwork\build\win\x64\Release\ProcessDuctwork.aip`
- Illustrator relaunched after the task ran

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

### Required Deployment Rule

- For panel/JSX/HTML/CSS changes only: copy the changed files to the CEP extension folder.
- For any C++ plugin or Emory workflow change: run `schtasks /run /tn "Reload Illustrator Ductwork"` instead of trying to copy the plugin into `Program Files` directly.
- Do not claim the plugin is deployed until the scheduled task finishes with `Last Result: 0` and the installed `ProcessDuctwork.aip` timestamp/hash matches the newly built file.

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
