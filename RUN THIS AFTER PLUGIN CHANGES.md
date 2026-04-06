# RUN THIS AFTER PLUGIN CHANGES

If a change touches the Emory/C++ plugin path, `cpp-plugin/`, or anything that affects `ProcessDuctwork.aip`, run this exact command:

```powershell
schtasks /run /tn "Reload Illustrator Ductwork"
```

Preferred wrapper:

```powershell
powershell -ExecutionPolicy Bypass -File ".\tools\reload-illustrator-ductwork.ps1"
```

Required verification before saying deployment is complete:

- `Reload Illustrator Ductwork` finished with `Last Result: 0`
- `C:\Program Files\Adobe\Adobe Illustrator 2024\Plug-ins\DuctworkMenu\ProcessDuctwork.aip` exists and was updated by the task
- The installed plugin hash matches `E:\Work\Work\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Process Ductwork\build\win\x64\Release\ProcessDuctwork.aip`
- Illustrator relaunched after the task finished

Direct `Program Files` copies are not the default deploy path for this project. Use the scheduled task.
