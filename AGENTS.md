# Magic Ductwork Panel - Project Instructions

## Critical

After finished with revisions deploy to C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\ automatically and prompt me for testing.

## Conversation Recovery

If a prior Codex conversation needs to be recovered, do not dump raw session JSONL into the chat or tool output. Raw session files contain huge encoded reasoning/tool blobs that can pollute and destabilize the next conversation. Extract only the useful handoff facts: user requests, assistant final messages, changed files, commits, deploy status, and command results.

Use `.Codex/recover-thread.ps1 codex://threads/<thread-id>` for recovery. Keep `.Codex/HANDOFF.md` updated after major milestones so a new thread can restart from clean facts instead of raw session data.

## Debug Logs Location

**IMPORTANT**: Runtime debug logs are stored at:
```
C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\Debug\
```

Log files are named with timestamps: `debug-YYYY-MM-DD_HH-MM-SS.log`

To read the latest debug log:
```powershell
powershell -Command "Get-ChildItem 'C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\Debug\' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { Get-Content $_.FullName -Tail 100 }"
```

After finished making changes deploy using the deployment path below

## Deployment Path

The extension is deployed to:
```
C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\
```

To deploy changes:
```powershell
powershell -Command "Copy-Item 'e:\Work\Work\Custom Sketchup, Illustrator and Photoshop Scripts and Extensions\Illustrator\Extensions\Magic-Ductwork-Panel\jsx\magic-final.jsx' 'C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\jsx\magic-final.jsx' -Force"
```

## Key Files

- `jsx/magic-final.jsx` - Main ductwork generation logic
- `index.html` - Panel UI
- `.Codex/feature-requests.html` - Feature request tracker
- `.Codex/issues.html` - Bug/issue tracker

## ExtendScript Notes

- Uses ES3 syntax (no let/const, no arrow functions, no template literals)
- `addDebug()` function writes to the debug log
- Path manipulation uses Illustrator's PathItem and pathPoints APIs
