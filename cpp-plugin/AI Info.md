# Build and Test

Use the full MSBuild path (plain `msbuild` is not on PATH in this shell):

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "E:\Work\Work\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Process Ductwork\src\ProcessDuctwork\ProcessDuctwork.vcxproj" /p:Configuration=Release /p:Platform=x64
```

Default deploy/test path for plugin changes:

```powershell
schtasks /run /tn "Reload Illustrator Ductwork"
```

Preferred wrapper from the panel repo root:

```powershell
powershell -ExecutionPolicy Bypass -File ".\tools\reload-illustrator-ductwork.ps1"
```

That scheduled task is the preferred path because it runs the elevated reload script, rebuilds the plugin, copies `ProcessDuctwork.aip` into `C:\Program Files\Adobe\Adobe Illustrator 2024\Plug-ins\DuctworkMenu\`, syncs the CEP panel, and launches Illustrator on the test file.

Do not stop at a successful build if the plugin path changed. Deployment is not complete until the scheduled task finishes with `Last Result: 0` and the installed plugin matches the built output.

Top-level reminder:

- `RUN THIS AFTER PLUGIN CHANGES.md`

Fallback manual admin command only if the task is unavailable:

```powershell
PowerShell -ExecutionPolicy Bypass -File "E:\Work\Work\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Process Ductwork\tools\reload-illustrator.ps1"
```

The companion interface for this extension is located at

E:\Work\Work\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Magic-Ductwork-Panel
