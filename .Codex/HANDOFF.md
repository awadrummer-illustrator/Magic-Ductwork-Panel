# Codex Handoff

Last updated: 2026-05-19

Use this file as the clean restart point when a Codex thread crashes or corrupts. Keep it short and human-readable. Do not paste raw session JSONL here.

## Current State

- Branch: `codex/normal-mode-fixes`
- Last known pushed panel commit before this recovery helper: `e00d7ac`
- Actual `Process Ductwork` repo was last verified clean at `35058d3`
- Deployed normal plugin hash was last verified as `47B1BFE65261033215EE21D1E28F63FC8F4453F12904C69ED549BCD0AFC61E34`
- Deployed `jsx/magic-final.jsx` was refreshed from the repo and verified at `EE85289D153A4BC9C0EF8A6D2CA6AB913D08F1EB8144FBA60A378E6C391D90ED`

## Recovery Workflow

Run the safe extractor from this repo root:

```powershell
.\.Codex\recover-thread.ps1 codex://threads/<thread-id>
```

For more context without dumping raw session data:

```powershell
.\.Codex\recover-thread.ps1 codex://threads/<thread-id> -AllAssistant -RecentCount 8
```

Only carry forward visible user requests, visible assistant conclusions, commits, deploy status, test results, and next actions.

## Latest Local Tooling Change

Added `.Codex/recover-thread.ps1`, which finds local Codex session logs by thread id and prints only safe, visible recovery information. It deliberately avoids raw JSONL, tool payloads, encoded reasoning, and giant command output.

## 2026-05-19 Emory Width Cascade Fix

- User reported Emory Mode width edits on green ductwork were too slow because edits cascaded through connected green runs and into blue ductwork, while prior branch sizing overrides kept scaling unexpectedly.
- Changed the Emory native width-apply path in sibling repo `Emory-Ductwork-Panel` so selected segment width edits stay local to selected runs and do not propagate into connected runs during rebuild.
- Updated Magic and Emory panel width-status text to explain that connected blue/branch runs stay unchanged.
- Built `EmoryDuctwork.aip` successfully with one pre-existing warning about an unused `sourceArt` parameter.
- Deployed panel JS to both CEP extension folders and deployed rebuilt `EmoryDuctwork.aip` to `C:\Program Files\Adobe\Adobe Illustrator 2024\Plug-ins\DuctworkMenu\` via the elevated reload task. Deployed plugin timestamp verified as `2026-05-19 04:38:12`.

## 2026-05-19 Emory Width Performance Pass

- Confirmed width calculations are native C++ through the Emory plugin; JSX only bridges the panel command into Illustrator.
- Panel log showed single width edits were still around `12s`, with native timing split roughly into `3s` state collection and `9s` generation/ordering.
- Updated Magic panel polling so Emory Mode no longer runs the slow ExtendScript selection hash poll, and width/stroke edits no longer schedule duplicate skip-ortho refreshes.
- Updated Emory panel width edit polling similarly.
- Changed native width apply to reuse the already-collected source states and network connections during the same rebuild, and to skip inherited-width pre-scan during protected width generation.
- Built `EmoryDuctwork.aip` successfully with the same pre-existing unused `sourceArt` warning.
- Deployed Magic panel JS, Emory panel JS, and the rebuilt plugin via reload task. Deployed plugin timestamp verified as `2026-05-19 04:53:53`.

## 2026-05-19 Emory Tee Width Fix

- User reported green T-junction openings did not match connected green segment widths after independent sizing.
- Fixed native endpoint-to-segment tee connector sizing so the branch mouth uses the connected branch segment width instead of clamping to the trunk width.
- Built `EmoryDuctwork.aip` successfully with the same pre-existing unused `sourceArt` warning.
- The normal reload helper did not update the plugin, so added `Emory-Ductwork-Panel\cpp-plugin\tools\deploy-emory-direct.ps1` and deployed through the existing elevated scheduled task forwarding hook.
- Deployed plugin timestamp verified as `2026-05-19 05:17:02`; Illustrator relaunched at `2026-05-19 05:17:57`.

## 2026-05-19 Emory Overlapping Tee Clamp

- User reported two overlapping T connectors should clamp to each other instead of overlapping.
- Added native connector clamp anchors for endpoint-to-segment and segment-intersection network connectors.
- Network connector arms now cap their length before the midpoint to the nearest neighboring connector point on the same source segment, preventing adjacent tee arms from growing through each other.
- Built `EmoryDuctwork.aip` successfully with the same pre-existing unused `sourceArt` warning.
- Deployed through `deploy-emory-direct.ps1`; deployed plugin timestamp verified as `2026-05-19 05:24:29`; Illustrator relaunched at `2026-05-19 05:24:43`.

## 2026-05-19 Emory Blue Start Sync

- User wanted the start of blue ductwork to match the green ductwork it directly connects to without bringing back the slow full blue cascade.
- Added native endpoint-only green-to-blue sync during Emory width apply. It updates only the connected blue start segment when the selected green/light-green source changed and the blue start width differs.
- Covered direct endpoint-to-endpoint green-to-blue connections and endpoint-to-segment cases where the selected green is the trunk and the blue starts as the branch endpoint.
- Deliberately does not sync orange runs, does not resize downstream blue segments, and does not overwrite a blue run that is itself selected.
- Updated Magic and Emory panel width-status text to say direct blue starts will match connected green.
- Built `EmoryDuctwork.aip` successfully with the same pre-existing unused `sourceArt` warning.
- Deployed Magic panel JS, Emory panel JS, and rebuilt plugin through `deploy-emory-direct.ps1`; deployed plugin timestamp verified as `2026-05-19 05:33:59`; Illustrator relaunched at `2026-05-19 05:36:53`.

## 2026-05-19 Emory Process-Time Blue Start Sync Fix

- User tested Process Ductwork with all green ductwork selected and the connected blue starts did not update.
- Root cause: prior blue-start sync was in the width-slider path, while Process Ductwork only generated selected green centerlines. The older generation inheritance could touch blue state in memory but did not persist or rebuild non-selected blue sources.
- Updated generation-time unit-pair handling so green/light-green is the authority for Blue Ductwork starts. Blue copies only the connected endpoint segment width; it does not cascade the width through the rest of the blue run.
- When generation touches a connected blue start, the blue source metadata is persisted, its old generated bodies are deleted, and that blue run is included in the same generation pass.
- Built `EmoryDuctwork.aip` successfully with the same pre-existing unused `sourceArt` warning.
- Deployed through `deploy-emory-direct.ps1`; deployed plugin timestamp verified as `2026-05-19 05:48:17`; Illustrator relaunched at `2026-05-19 05:48:32`.

## 2026-05-19 Emory Move Panel Generated Endpoint Rebuild

- User wanted Direct Selection of source centerline endpoints or generated blue Emory ductwork endpoints to work with the move panel as if the real centerline endpoint was selected.
- Added native selection resolving so selected generated blue Emory segment/guide endpoints map back to the owning blue centerline terminal endpoint and return the source run ID to JSX.
- Added native `rebuild-emory-sources` panel action that refreshes final segment omit metadata, deletes only the affected generated blue bodies, and rebuilds only those blue source runs.
- Updated `jsx/panel-bridge.jsx` so generated Emory bodies use resolved source endpoints instead of selected polygon corners, place/replace the selected register/unit anchor there, and queue a targeted blue rebuild.
- Updated `cpp-plugin\tools\deploy-emory-direct.ps1` so future direct deploys copy the JSX bridge as well as the panel JS.
- Built `EmoryDuctwork.aip` successfully with existing warnings only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 06:07:45`; JSX bridge copied and Illustrator relaunched at `2026-05-19 06:08:20`.

## 2026-05-19 Magic Move Panel Rebuild Bridge Fix

- User checked the move log after testing and the register was added but blue Emory ductwork was not rebuilt.
- Root cause: the active UI was `Magic-Ductwork-Panel`, whose move bridge still called the old `ProcessDuctwork` anchor detector and never called the new Emory rebuild action. The missing `Emory rebuilt sources` line in `Magic-Ductwork-Panel\move-debug.log` confirmed this.
- Updated Magic `jsx/panel-bridge.jsx` to prefer `EmoryDuctwork` `get-selected-anchors`, preserve source IDs on matched source centerline endpoints, resolve generated blue body endpoints, and call `rebuild-emory-sources` after placing/replacing the register anchor.
- Deployed updated `jsx/panel-bridge.jsx` to `C:\Users\Chris\AppData\Roaming\Adobe\CEP\extensions\Magic-Ductwork-Panel\jsx\panel-bridge.jsx`; deployed timestamp verified as `2026-05-19 06:21:40`. Also refreshed the Emory panel bridge copy at `2026-05-19 06:21:47`.
- Tightened source endpoint handling so retrying after a failed placement can replace the existing part and still queue the blue source rebuild instead of skipping because a register already exists there.
- User needs to reload the Magic panel or Illustrator before retesting because the move panel caches the bridge after first load.

## 2026-05-19 Mixed Selection Endpoint Fix

- User checked the log after selecting existing square registers plus blue ductwork endpoints and changing to rectangular registers. Only some selected square registers were replaced; selected Blue Ductwork endpoints were ignored.
- Root cause: `prioritizePartReplacement` mode skipped every non-part item whenever any existing part was selected. The log showed every selected Blue Ductwork path as `Ignored in mixed-selection replacement mode`.
- Updated Magic and Emory `jsx/panel-bridge.jsx` so mixed-selection replacement still processes ductwork color `PathItem`s. Non-part/non-ductwork items remain ignored, but blue endpoints now reach the endpoint extraction/rebuild logic.
- Deployed updated Magic bridge to CEP timestamp `2026-05-19 06:29:10`; deployed updated Emory bridge timestamp `2026-05-19 06:29:20`.
- User needs to reload the Magic panel or Illustrator before retesting because the bridge is cached after load.

## 2026-05-19 Emory Green-to-Blue Tail Anchor Skip

- User reported the older internal-anchor logic that adds anchors partway down single-segment branches should not apply to green ductwork branches that connect to blue ductwork.
- Root cause: `AddRegisterTailAnchors` only had a partial green-to-blue skip for endpoint-to-endpoint unit-pair transitions inside the selected path set. It missed green/light-green endpoints landing on blue segments, especially when blue ductwork was not selected.
- Updated Emory native `ProcessDuctworkPlugin.cpp` to collect unselected Blue Ductwork centerlines as context and skip register-tail/internal-anchor insertion for Green Ductwork or Light Green Ductwork when either endpoint touches a blue endpoint or lands on a blue segment.
- Built `EmoryDuctwork.aip` successfully with existing warnings only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 06:49:19`; Illustrator relaunched at `2026-05-19 06:49:36`.

## 2026-05-19 Emory Branch Drop Apply Fix

- User reported selecting generated Emory ductwork and entering a new branch drop percentage did not change the visible taper.
- Root cause: the Magic panel branch drop field only saved the value for future Process Emory runs. It did not call any native apply/rebuild action for the selected generated Emory ductwork.
- Added native `apply-emory-branch-taper` action. It resolves selected generated Emory pieces back to their source run IDs, applies the new branch drop percentage, recomputes selected branch widths from the parent trunk, deletes old generated bodies, rebuilds the selected runs, and reselects the same generated segment indices when possible.
- Added `MDUX_cppApplySelectedEmoryBranchTaperReduction(...)` bridge wrapper in Magic and Emory bridge files.
- Updated Magic `js/panel.js` so changing the branch drop slider, committing the input, or pressing Enter applies the value to the current selected generated Emory branch instead of only storing it.
- Built `EmoryDuctwork.aip` successfully with existing warnings only.
- Deployed Magic `js/panel.js` and `jsx/panel-bridge.jsx` to CEP; deployed rebuilt Emory plugin and Emory bridge through `deploy-emory-direct.ps1`. Scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 07:09:05`; Illustrator relaunched at `2026-05-19 07:09:51`.

## 2026-05-19 Emory Branch Drop Taper Math Fix

- User reported the new branch drop apply action rebuilt selected generated ductwork but the taper still did not change according to the new percentage.
- Root cause: branch drop changed the branch inherited root width, but branch straight-chain taper still used the hard-coded `kStraightTaperRatio = 0.8`. This meant the visible per-segment taper proportion stayed the same even when the branch drop percentage changed.
- Parameterized the straight-chain taper helpers and changed inherited branch width generation to use `gBranchInheritedWidthRatio` as the taper step for branch-retapered runs.
- Example expectation: a 50% branch drop now uses a 0.5 taper step for inherited branch chain widths instead of preserving the old 0.8 step.
- Built `EmoryDuctwork.aip` successfully with one existing warning only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 07:18:20`; Illustrator relaunched at `2026-05-19 07:18:35`.

## 2026-05-19 Emory Reset Stored Widths Button

- User asked for an explicit button under the Emory sliders to reset selected ductwork widths back to automatic taper behavior, because changing Branch Drop already did that but was easy to forget.
- Added a `Reset Stored Widths` button under the Emory width/stroke/branch-drop sliders in Magic `index.html`; also added it to the standalone Emory panel.
- Added native `reset-emory-stored-widths` action. It resolves selected generated Emory pieces/source centerlines to source run IDs, recomputes selected non-branch runs from their default taper, recomputes selected branch runs from the current Branch Drop percentage and parent trunk width, persists the recomputed widths, rebuilds only selected touched runs, and reselects selected generated segment indices when possible.
- Added `MDUX_cppResetSelectedEmoryStoredWidths(...)` bridge wrappers in Magic and Emory bridge files.
- Updated Magic `js/panel.js` so the button uses the current Branch Drop field value; standalone Emory calls the same action with the default 25% branch drop because that older UI has no Branch Drop control.
- Updated `cpp-plugin\tools\deploy-emory-direct.ps1` so future direct Emory deploys copy `index.html` too, not only JS/JSX.
- Built `EmoryDuctwork.aip` successfully with existing warnings only.
- Deployed Magic `index.html`, `js/panel.js`, and `jsx/panel-bridge.jsx` to CEP. Deployed standalone Emory `index.html`, `js/panel.js`, and `jsx/panel-bridge.jsx` to CEP. Deployed rebuilt Emory plugin through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 07:47:00`; latest Illustrator relaunch completed at `2026-05-19 07:49:25`.

## 2026-05-19 Emory Selected Blue Segment Branch Drop Cascade Fix

- User reported a particular Blue Ductwork run still did not respect a smaller Branch Drop percentage, and Reset Stored Widths did not fix it.
- Latest native log showed `branchTaperReductionPercent=5` and inheritance factor `0.95`, but problem source `emory-1779180935024-71` generated widths `[13.44,6.4,5.12,4.096]`. That proved part of the run was still using old stored/default `0.8` taper values instead of the selected `0.95` factor.
- User clarified expected behavior: selecting the first blue generated segment that connects to green and changing Branch Drop should cascade through that blue branch/run without selecting every blue segment.
- Added selected-source retapering to `apply-emory-branch-taper`: after normal parent/child width inheritance, the selected source run is retapered from the selected generated segment when one is selected, using the current Branch Drop factor. The selected segment becomes the stored start segment for future rebuilds.
- Updated Reset Stored Widths to use the same selected-segment retaper path instead of retapering non-branch/unit-connected runs with the old default ratio.
- Updated blue/orange unit-endpoint taper normalization to use the current Branch Drop factor instead of hard-coded `kStraightTaperRatio`.
- Built `EmoryDuctwork.aip` successfully with one existing warning only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 08:00:21`; Illustrator relaunched at `2026-05-19 08:00:36`.

## 2026-05-19 Emory Reset Connected Branch Widths Fix

- User reported Reset Stored Widths fixed the selected main blue run, but branches off that blue trunk stayed too thick instead of inheriting from the trunk segment they connect to.
- Latest native log showed selected source `emory-1779180935025-113` retapering correctly to widths `[15,11.25,8.4375]`, with no follow-up reset/rebuild entries for connected branch sources.
- Added `ExpandAffectedSourcesToConnectedBranches(...)` so Branch Drop apply and Reset Stored Widths walk endpoint-to-segment, unit-feeder, and same-layer intersection branch connections outward from the selected source IDs.
- Updated both selected Branch Drop apply and Reset Stored Widths to retaper the selected source, expand to connected branch sources, reapply inherited branch widths from the updated trunk widths, persist every affected source, and rebuild every touched selected/connected source.
- Added log line `Emory affected branch expansion added=... affectedSources=...` for this path.
- Built `EmoryDuctwork.aip` successfully with one existing warning only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 08:15:52`; deploy completed at `2026-05-19 08:16:05`.

## 2026-05-19 Emory Connected Branch Flat-Taper Followup

- User retested and reported Reset Stored Widths still appeared to do nothing, while changing Branch Drop rebuilt but left the branch pieces visibly too thick.
- Latest native log after the `08:16` deploy showed Branch Drop did reach C++ and rebuilt sources `106` through `113`, but connected branch sources such as `108`, `109`, `111`, and `112` rebuilt with flat widths like `[11.25,11.25,11.25]`.
- Root cause: inherited branch reset anchored the branch at the parent connection, but then used the old straight-chain-only taper. Bent branch centerlines therefore kept a flat inherited root width across the visible branch segment.
- Added native entry logs for both Branch Drop apply and Reset Stored Widths:
  - `Emory branch-drop apply requested reduction=... factor=... selectedSources=... selectedSegments=...`
  - `Emory reset stored widths requested reduction=... factor=... selectedSources=... selectedSegments=...`
- Added `ApplyBranchDropTaperAcrossSegments(...)` and changed selected retaper, endpoint branch inheritance, same-layer intersection branch inheritance, and unit-feeder downstream inheritance to taper outward from the actual connection segment across the branch, including bends.
- Added per-branch inheritance logs showing trunk source, branch source, connection segment, root width, factor, and final width list.
- Built `EmoryDuctwork.aip` successfully with one existing warning only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 08:25:14`; Illustrator relaunched at `2026-05-19 08:25:38`.

## 2026-05-19 Emory Ignore Endpoint Ortho Fix

- User reported one line in the last processed selection was not being ortho'd. It had previously ended at a register while No Ortho Final was on, but was later changed to an ignore marker, so it should have been eligible for ortho again.
- Root cause: native `DuctworkOrtho::ApplyToPaths(...)` skipped final branch segments based only on an unconnected branch endpoint. It did not know that an endpoint near an `Ignore`/`Ignored` anchor should no longer be treated like a final register tail.
- Updated `ProcessDuctworkOrtho.cpp` to collect point anchors from `Ignore`/`Ignored` layers before applying No Ortho Final, mark selected path endpoints within 10 pt of those anchors, and allow those ignored endpoints to be ortho'd instead of skipped.
- Removed two stale unused variables in the same file so the native build is clean.
- Built `EmoryDuctwork.aip` successfully with 0 warnings and 0 errors.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 08:40:08`; deploy completed and Illustrator relaunched at `2026-05-19 08:40:24`.

## 2026-05-19 Emory Internal-Vertex Blue Junction Fix

- User reported two Blue Ductwork lines were not being treated as one connected run and were not creating a T junction.
- Latest native log for the selected pair showed `ductworkPaths=2 connections=0`, so the connector builder never saw the two selected blue centerlines as connected.
- Root cause: `DuctworkConnections::FindConnections(...)` discarded segment intersections near segment endpoints. That avoided false endpoint connectors, but it also discarded real same-color junctions when the connection point landed on an internal polyline vertex created by ortho/register-tail splitting. Whole-path endpoint-to-segment detection did not cover internal vertices, so the connection disappeared entirely.
- Updated `ProcessDuctworkConnections.cpp` so internal-vertex same-color intersections are kept, while true terminal endpoint intersections are still skipped.
- Added a junction-arm fallback in `ProcessDuctworkGeometry.cpp`: if the normal endpoint-to-segment or segment-intersection connector builder cannot build a connector, it now gathers all source segment arms around the junction point and builds a tee/cross connector from those arms.
- Added log line `Emory junction connector fallback sourceA=... sourceB=... style=... arms=... point=...` for retesting this case.
- Built `EmoryDuctwork.aip` successfully with one pre-existing warning (`sourceArt` unreferenced).
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 09:17:38`; deploy completed and Illustrator relaunched at `2026-05-19 09:17:56`.

## 2026-05-19 Emory Blue Unit Feeder Internal Anchor Removal

- User asked why the last run added an internal anchor near the unit.
- Latest native log showed `Register tail path splits inserted=1` before ortho, and the selected Blue Ductwork unit feeder source `emory-1779180935024-58` generated with `segments=2`, proving the source centerline had been split by the old unit-pair branch taper normalizer.
- Root cause: `AddRegisterTailAnchors(...)` still called the unit-pair branch taper normalizer for Blue/Orange runs where one endpoint is near a unit/thermostat pair and the opposite endpoint connects into a same-color trunk. That legacy code inserted a synthetic point at 65% from the unit endpoint to the trunk endpoint.
- Removed the insertion path. Unit-pair Blue/Orange feeders that connect to a same-color trunk now log `Unit-pair branch taper anchor skipped ...` and continue without adding a synthetic point.
- Added a narrow cleanup: if the selected source already has the old straight-line synthetic point near that 65% position, it collapses the centerline back to the two real endpoints and logs `removedSynthetic=1`. Real drawn bends are left alone.
- Built `EmoryDuctwork.aip` successfully with existing warning noise only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 09:29:39`; deploy completed and Illustrator relaunched at `2026-05-19 09:30:00`.

## 2026-05-19 Emory Blue Register Tail Anchor Correction

- User clarified that the partial internal anchor on Blue Ductwork is still required for real register branches; only the unit/thermostat feeder case should be blocked.
- Reworked `AddRegisterTailAnchors(...)` so Blue Ductwork register branches can still get the 65% register-tail anchor when one side is a trunk connection and the other side is the register/open terminal.
- Added unit/thermostat endpoint detection using existing `Units`, `Thermostats`, and `Thermostat Lines` layer points. If the would-be blue register endpoint is actually near a unit/thermostat endpoint, the anchor is skipped and logs `Blue unit tail anchor skipped ...`.
- Narrowed old-anchor cleanup so it removes a straight 3-point Blue Ductwork source only when exactly one endpoint is near a unit/thermostat and neither endpoint is near a register. This preserves real register-tail anchors.
- Added geometry-side sanitizing so generated Emory bodies and final-segment metadata also ignore that stale unit-feeder synthetic point if it already exists on the source/backup, while preserving register branch splits.
- Built `EmoryDuctwork.aip` successfully with existing warning noise only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deploy completed and Illustrator relaunched at `2026-05-19 09:43:28`.

## 2026-05-19 Emory Do Not Remove Existing Internal Anchors

- User clarified the unit-feeder fix must not delete or ignore existing internal anchors; it should only prevent the bad anchors from being newly inserted.
- Removed the source mutation cleanup that collapsed existing 3-point Blue Ductwork unit feeders back to two endpoints.
- Removed the geometry-side sanitizing that ignored existing unit-feeder internal points during generated body/final-thickness calculation.
- Current behavior: Blue register branches can still receive the required 65% register-tail anchor. Blue unit/thermostat feeders skip the new tail-anchor insertion when the would-be terminal side is near a unit/thermostat. Existing internal anchors are left alone.
- Built `EmoryDuctwork.aip` successfully with existing warning noise only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin timestamp verified as `2026-05-19 09:49:10`; deploy completed and Illustrator relaunched at `2026-05-19 09:49:26`.

## 2026-05-19 Emory Unit Handoff Width Lock

- User identified the core usability problem: green and blue ductwork segments that meet at generated unit handoffs need to be selectable as a pair, set to a width together, and then stay at that width. Blue width changes should not cascade back up and rewrite the green handoff branch.
- Added native action `select-emory-unit-pair-segments`. It looks at the current selected generated/centerline Emory source IDs, finds green/blue unit-pair connections using the existing unit attachment resolver, and selects the generated segment on both sides of each matching handoff.
- Added panel button `Select Unit Handoff Segments` under the existing Emory selection buttons in both the Emory CEP panel and the Magic Emory-mode CEP panel. Flow: select either side of the handoff, click this button, then set Width.
- Updated Width apply so when both sides of a unit handoff are selected, those selected segments are set to the exact requested width instead of being proportionally scaled. The same segments are marked as cascade stops so later taper/branch recalculation does not push blue-side width changes back into green-side handoff widths.
- Updated backup centerline sync so those cascade-stop locks persist when the visible generated source is backed by a hidden centerline fragment.
- Built `EmoryDuctwork.aip` successfully with existing warnings only.
- Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; deployed plugin, Emory CEP files, and Magic CEP files verified. Illustrator was relaunched again after Magic CEP copy; final reload completed at `2026-05-19 11:46:48`.

## 2026-05-19 Emory Unit Handoff Button Disabled Fix

- User reported `Select Unit Handoff Segments` did nothing.
- Checked latest native and panel logs. There was no native `select-emory-unit-pair-segments` call and no panel evalScript entry for `MDUX_cppSelectSelectedEmoryUnitPairSegments()`, so the click was not reaching the plugin.
- Root cause: the new selector was being disabled by `setEmoryControlsEnabled(false)` whenever the current selection did not qualify as width-editable. That made the discovery/selection tool unavailable exactly when it was needed.
- Updated both Emory CEP and Magic Emory-mode CEP panels so `selectEmoryUnitPairsBtn` stays enabled independently of Width controls.
- Added panel log line `[EMORY] Select Unit Handoff Segments click` and native log lines beginning `Emory select unit handoff...` to make the next failure mode visible immediately.
- Rebuilt `EmoryDuctwork.aip` successfully with one existing warning (`sourceArt` unreferenced).
- Copied Magic CEP panel files, deployed Emory plugin/panel through `deploy-emory-direct.ps1`, scheduled task result `0`, and verified installed timestamps. Final deploy completed at `2026-05-19 11:59:06`.

## 2026-05-19 Emory Unit Handoff Selection Source-ID Fallback

- User retested after selecting first and reported the handoff button still did nothing.
- Fresh native log showed the action was firing repeatedly but `CollectSelectedEmorySourceIds(...)` returned zero source IDs: `Emory select unit handoff segments requested` followed by `no selected source IDs`.
- Root cause: the source-ID collector only used the normal selected-path helper. Some Illustrator selections, especially Direct Select anchor/point selections or generated compound/group child selections, do not expose the source metadata on the exact selected path handle.
- Expanded `CollectSelectedEmorySourceIds(...)` to gather Emory source IDs from raw selected art, selected path descendants, parent art containers, existing generated/source metadata, and direct-selected path-point states across line-layer paths.
- Built `EmoryDuctwork.aip` successfully with one existing warning (`sourceArt` unreferenced).
- Copied Magic CEP panel files, deployed Emory plugin/panel through `deploy-emory-direct.ps1`, scheduled task result `0`, deployed plugin timestamp `2026-05-19 12:05:22`, deploy completed at `2026-05-19 12:05:41`.

## 2026-05-19 Emory Unit Handoff Endpoint Pair Fallback

- User checked the log again after the source-ID fallback. Native logs showed the button/action was firing and found 189 selected Emory sources, but still selected nothing: `no pair selected connections=151 unitAttachments=74`.
- Root cause: the unit handoff selector depended on `CollectEmoryNetworkConnections(...)` plus `ResolveUnitPairEndpointConnection(...)`. In this document the green/blue unit handoff endpoints were not present as unit-pair connections in that network list, even though the source endpoints themselves touch where units are created.
- Updated `SelectSelectedEmoryUnitPairSegments(...)` with a direct endpoint fallback. It now scans paired green/light-orange transition runs against their blue/orange main runs and treats touching endpoints within 12 pt as unit handoff pairs, independent of the network connection list.
- Updated `CollectSelectedUnitPairHandoffSegments(...)` with the same endpoint fallback so the follow-up Width action can lock those selected handoff pairs as cascade stops. This avoids fixing the selection button but leaving width-lock behavior broken.
- Build passed with one existing warning (`sourceArt` unreferenced).
- Copied Magic CEP panel files and deployed Emory plugin/panel through `deploy-emory-direct.ps1`; scheduled task result `0`; final reload completed at `2026-05-19 12:15:11`.

## 2026-05-19 Emory Same-Layer Internal Marker Connector Suppression

- User reported the last ductwork processing ignored an internal anchor at the intersection of two Blue Ductwork lines, where the internal anchor should mean "do not create a T/cross junction here."
- Latest native log showed `sameLayerMarkerCount=2` on source `emory-1778272818088-2561`, followed by `Emory junction connector fallback ... style=cross ... point=[133,1561]` and another at `[133,1442.59]`. So the marker was detected for marker-span/taper handling, but the network connector fallback still built visible cross connectors at those marker points.
- Added `ShouldSuppressNetworkConnectorForSameLayerMarker(...)` in `ProcessDuctworkGeometry.cpp`. Segment-intersection network connectors are now skipped when either same-layer source has an intentional same-layer internal marker vertex at the connection point.
- Also excluded those marker connections from network connector clamp anchors, so an intentional "no junction" marker does not shorten nearby connector arms.
- Added log line `Emory network connector skipped same-layer marker ...` for retesting.
- Build passed with one existing warning (`sourceArt` unreferenced). Deployed through `deploy-emory-direct.ps1`; scheduled task result `0`; final reload completed at `2026-05-19 12:33:13`.
- User clarified that same-layer internal intersection anchors also must not create taper at the marker. Root cause: marker width normalization handled each marker independently, so adjacent markers could still leave an earlier marker with mismatched segment widths after a later marker raised the shared middle segment.
- Updated `NormalizeSameLayerIntersectionMarkerWidths(...)` so consecutive same-layer marker vertices are normalized as a full span. In the logged case, the marked span should normalize the middle run to one width instead of leaving a taper at either marker.
- Rebuilt successfully with one existing warning (`sourceArt` unreferenced). Deployed via the prompt-preserving scheduled-task/request-file path; user confirmed they received the permission popup. Task result `0`; plugin copy timestamp `2026-05-19 12:35:37`; Illustrator relaunched at `2026-05-19 12:40:16`.

## Operator Preference: Illustrator Reload Permission

- Do not start `Reload Illustrator Ductwork` or otherwise restart Illustrator without explicit user permission first.
- Build and prepare deploy files as needed, but pause before any Illustrator reload/restart and ask the user whether to reload now.
- User specifically wants the permission popup/confirmation flow respected; do not bypass it with direct scheduled-task starts.
- Root cause of the bypass: `Reload Illustrator Ductwork` points at the Process Ductwork reload wrapper, and that wrapper forwards to the path in `%TEMP%\ductwork-reload-request.txt` before showing its own popup. When the request file points at Emory's `cpp-plugin\tools\deploy-emory-direct.ps1`, the direct script must show its own confirmation.
- Updated Emory `cpp-plugin\tools\deploy-emory-direct.ps1` to show a topmost Yes/No popup before stopping/restarting Illustrator. Syntax check passed. Do not run it without user confirmation.

## 2026-05-20 Emory Register Tail Anchor and No-Ortho Fix

- User reported that the required internal anchors on single-line Blue Ductwork register branches appeared to be gone, despite explicitly needing to preserve them for register cases.
- Investigation showed `AddRegisterTailAnchors(...)` still existed, but selected-only processing could hide the same-layer trunk/context connection from the anchor inserter. Added unselected Blue Ductwork context paths into the connection scan so a selected register branch can still see the trunk it connects to.
- Kept the unit/thermostat guard narrow: Blue unit/thermostat feeder lines still skip new synthetic tail anchors, but Blue register branches can still receive the 65% internal anchor.
- User also clarified that `No Ortho Final` must not stop orthoing a single-line Blue Ductwork branch before that internal register anchor gets added.
- Updated `ProcessDuctworkOrtho.cpp` so `No Ortho Final` only skips the final register tail when the path already has more than one segment (`count > 2`). A raw one-segment register branch remains eligible for ortho first, then the register-tail anchor can be added.
- Built `EmoryDuctwork.aip` successfully with 0 warnings and 0 errors.
- Deployed through the popup-preserving scheduled task. Deploy log shows `Direct Emory deploy complete` at `2026-05-20 02:15:52`.

## 2026-05-20 Emory Width/Branch Drop Start Segment Rule

- User reported setting the last selected Blue Ductwork width to `12` snapped back to `5.32446`, and some blue branches shrank far more than the configured `15%` drop.
- Latest native log showed the branch-drop pass selected 188 sources and logged `factor=0.85`, but sources such as `emory-1779256948269-41` were retapered from `startSegment=0` even though their unit-side default was segment `3`. That cascaded `12 -> 10.2 -> 8.67 -> 7.3695 -> 6.264 -> 5.32446`.
- User clarified the rule: taper should occur from the start segment; by default, the start segment should be where the units are, unless the user explicitly marks a different segment as the start.
- Changed `ReadStartSegmentIndex(...)` so an explicitly stored start segment is honored first, and otherwise the default resolver chooses the unit/paired-unit side. Previously directional defaults could override stored starts.
- Changed `HasExplicitStartSegmentIndex(...)` to mean the start metadata actually exists, instead of returning false whenever a unit-side directional default was available.
- Changed Branch Drop and Reset Stored Widths so they use each run's current start segment. They no longer treat arbitrary selected generated segments as new taper anchors during broad selections.
- During width-apply rebuilds, protected selected sources skip the blue/orange unit endpoint normalizer, and the normalizer is also skipped when a run has an explicit start segment. This prevents a manual width like `12` from being lowered back to inherited taper values during the same rebuild.
- Build passed with one existing warning (`sourceArt` unused). Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 02:47:38`.

## 2026-05-20 Thermostat Endpoint Snap Tolerance

- User asked to increase the distance a thermostat line endpoint can be from Blue or Green Ductwork endpoints before it snaps to that endpoint and suppresses thermostat creation, by about 65%.
- Updated native `SnapThermostatEndpoints(...)` in `ProcessDuctworkPlugin.cpp`: snap target endpoints now include `Blue Ductwork`, `Green Ductwork`, and `Light Green Ductwork` source endpoints instead of only Blue. The snap tolerance changed from `12.0` to `20.0` points, with log line `Thermostat endpoints snapped tolerance=20`.
- Updated legacy Magic JSX fallback `THERMOSTAT_JUNCTION_DIST` in `jsx/magic-final.jsx` from `6` to `10`, matching a 65% bump (`6 * 1.65 = 9.9` rounded).
- Built `EmoryDuctwork.aip` successfully. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 02:59:14`, installed plugin timestamp `2026-05-20 02:58:41`.

## 2026-05-20 Thermostat Snap Document Context and Cleanup Logging

- User reported thermostat endpoints within the new 20 pt tolerance still created thermostats instead of snapping. Latest log had no `Thermostat endpoints snapped` line, while `Parts: supplemented Thermostats paths from document count=33` showed thermostat lines were being reintroduced later by part creation.
- Root cause: `BuildProcessPathEntry(...)` filters thermostat lines out of the selected processing set, so the snap tolerance change was real but the snap pass often never saw the thermostat line paths. Part creation then scanned document thermostat lines and created thermostat parts from their unsnapped endpoints.
- Added `BuildThermostatSnapPathEntry(...)` and moved thermostat snapping into a separate document-context pass before ortho and before parts. It scans all line-layer paths, accepts open `Thermostat Lines` directly, accepts real Blue/Green/Light Green centerline candidates, and snaps thermostat endpoints to duct endpoints within `20.0`.
- New log line: `Thermostat endpoint snap checked lines=<n> ductEndpoints=<n> snapped=<n> tolerance=20.000000`, plus `Thermostat endpoints snapped count=<n> tolerance=20` when anything moves.
- User also asked whether process ductwork was creating duplicate geometry instead of replacing it. The log showed older runs with `Emory bodies deleted=0` followed by `created=573`, then the latest broad run cleaned up `1107` before creating `579`, meaning stacked old geometry was present and got removed in that run.
- Added cleanup logging `Emory cleanup sourceIds=<n>` and included source IDs from selected generated Emory art in the cleanup list via `AppendSelectedGeneratedEmoryCleanupIds(...)`. This should help select-all/broad selections delete old selected generated bodies even when the selected processing centerline set does not fully account for them.
- Built successfully with existing warnings only. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 03:12:41`, installed plugin timestamp `2026-05-20 03:12:17`.

## 2026-05-20 Reset Stored Widths Selected-Branch Inheritance

- User asked why a section near `11.95/12` wide had a connected branch around `4.95`, far more than the expected branch-drop percentage.
- Latest native log showed the active drop setting was `10%` (`factor=0.9`), so the large mismatch was not honest taper math. The tiny branch was retaining an old stored width and never logging parent inheritance from the nearby `12` trunk.
- Root cause: `ApplyInheritedBranchWidths(...)` skipped `selectedSeed` targets. That protection is correct for manual Set Width, but wrong for Branch Drop and Reset Stored Widths because those commands are explicitly supposed to make selected/connected branches follow parent inheritance again.
- Updated `ApplyInheritedBranchWidths(...)` and `ApplyDirectUnitPairEndpointWidthSync(...)` with an `allowSelectedSeedTargets` flag. Normal generation/manual width paths pass `false`, preserving selected manual-width protection. Branch Drop and Reset Stored Widths pass `true`, so selected stale branches can be overwritten by inherited parent widths and the current drop percentage.
- Added `allowSelectedSeed=<0|1>` to the inheritance log line for retesting.
- Verified Magic panel wiring: `Reset Stored Widths` in Magic reads the current Emory branch-drop percentage and calls `MDUX_cppResetSelectedEmoryStoredWidths(value)`. The older standalone Emory panel still sends a fixed `25` because it does not have the same drop control.
- Built `EmoryDuctwork.aip` successfully with one existing warning (`sourceArt` unused). Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 03:25:18`.

## 2026-05-20 Same-Layer Marker Width Sync

- User reported the same branch still did not respect the ductwork drop-off percentage.
- Fresh native log showed the specific missed case: `sourceId=emory-1779258110891-624` stayed around `3.9304`, while intersecting `sourceId=emory-1779258110891-626` had a same-layer internal marker span at `8.67`. The connector was correctly skipped (`Emory network connector skipped same-layer marker...`), but width inheritance skipped the connection too because `ResolveSameLayerSegmentIntersectionBranch(...)` rejects internal-vertex intersections.
- Added a same-layer marker width sync path. Internal-marker intersections still do not create a visible T connector, but the unmarked crossing run now takes the marked segment width at the crossing, then applies the branch-drop taper away from that segment.
- Added expansion of affected source IDs across same-layer marker intersections so Reset Stored Widths / Branch Drop can rebuild and persist the connected crossing run even when only the marked run started as affected.
- New log line for retesting: `Emory same-layer marker width synced source=<marked source> target=<crossing source> ...`.
- Built `EmoryDuctwork.aip` successfully with one existing warning (`sourceArt` unused). Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 03:34:33`, installed plugin timestamp `2026-05-20 03:34:05`.
- User checked the log and noted the panel showed the selected segment start ductwork width as `11.95`. The 03:36 log showed the first marker sync patch fired, but still copied raw segment `0` from `626` as `7.3695`: `Emory same-layer marker width synced source=...626 target=...624 sourceSegment=0 targetSegment=1 rootWidth=7.3695`. Later generation normalized the marker span on `626` to `8.67`, so `624` remained one drop too small.
- Updated same-layer marker sync to resolve the marker span width the same way generation does: collect same-layer marker vertices, find the contiguous marker span, use the max width across the span, then sync that width to the crossing run before applying taper away from the marker segment.
- Built successfully again with the same existing warning. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 03:39:19`, installed plugin timestamp `2026-05-20 03:38:44`.

## 2026-05-20 Blue Branch Double-Drop Width Fix

- User reported a blue branch connected to the blue line coming off green/unit handoff was still tiny, far more than the selected branch-drop percentage.
- Latest native log showed the bad pattern on sources such as `emory-1779258110892-652 -> 650/651`: trunk `652` retapered from start segment `1` to widths `[6.88,8,6.88,5.9168,...]`, then child branches connected to segment `2` inherited `5.9168` because inheritance applied the branch drop again to the already-dropped segment width.
- Changed endpoint-to-segment and same-layer intersection branch inheritance to resolve a blue/orange trunk's upstream split width first. If the branch connects to a segment downstream of the trunk start, the child inherits from the adjacent segment toward the start, then applies the branch-drop percentage once. Cascade-stop/locked trunk segments still use the locked segment width directly.
- Added log fields `parentBasisSegment` and `parentBasisWidth` to `Emory endpoint branch inherited...` and `Emory intersection branch inherited...` for retesting.
- Built `EmoryDuctwork.aip` successfully with the existing `sourceArt` warning only. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 03:51:25`; installed plugin timestamp verified at `C:\Program Files\Adobe\Adobe Illustrator 2024\Plug-ins\DuctworkMenu\EmoryDuctwork.aip` as `2026-05-20 03:50:58`.

## 2026-05-20 Selected Blue Feeder Width Cascade Fix

- User checked the log after setting source `emory-1779258110889-554` to `12`; the native log confirmed `554` was `12`, but connected blue run `555` stayed at `4.913`, and its register branches stayed at `4.17605`.
- Root cause: manual width apply had been made local-only for performance, so selected blue feeder/start segments did not push their width into the connected downstream blue run. A cascade-stop/lock on the selected feeder also blocked the unit-feeder handoff path from applying to the next run.
- Changed `ApplyContinuationWidthFromUnitFeeder(...)` so a cascade stop on the source feeder no longer blocks the handoff; only a cascade stop on the downstream target segment blocks overwriting that target.
- Changed manual width apply so non-terminal selected Blue/Orange runs expand only through connected downstream branches and run `ApplyInheritedBranchWidths(...)` from those selected blue sources. This keeps the slider from broadly cascading green-to-blue, but lets a selected blue feeder/root update the connected blue trunk and its child branches using the current branch-drop percentage.
- New log line for retesting: `Emory width-apply cascaded selected blue feeder/run widths expanded=<n> inherited=<n>`.
- Built `EmoryDuctwork.aip` successfully with the existing `sourceArt` warning only. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 04:00:59`.

## 2026-05-20 Unit-Feeder To Blue Trunk Drop Fix

- User clarified the previous fix made the blue trunk that all register branches connect to grow, but not by the drop-off percentage. The branches off that trunk changed correctly when branch drop changed, but the middle trunk between the unit feeder and register branches was using the unit-feeder width at 100%.
- Latest log example: with factor `0.86`, `555` stayed `11.95` and branches `549-553` became `10.277`; with factor `0.5`, `555` still stayed `11.95` and branches became `5.975`. Expected model is `554` unit feeder stays explicit, `555 = 554 * factor`, and `549-553 = 555 * factor`.
- Changed `ApplyContinuationWidthFromUnitFeeder(...)` so same-layer Blue/Orange unit-feeder-to-segment handoffs apply the current branch-drop factor to the downstream trunk root (`rootWidth = feederWidth * gBranchInheritedWidthRatio`) instead of copying the feeder width directly.
- Changed `ApplyInheritedBranchWidths(...)` so that unit-feeder downstream recalculation runs when either the feeder or downstream trunk is affected. This lets Branch Drop / Reset Stored Widths fix the middle trunk even if the unit feeder itself was not changed in that command.
- The unit-feeder log now includes `feederWidth`, `factor`, and `rootWidth`.
- Built successfully with the existing `sourceArt` warning only. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 04:08:05`; installed plugin timestamp is `2026-05-20 04:07:34`.

## 2026-05-20 Branch Drop No Longer Drops At Every Turn

- User asked whether the blue branch drop was applied at every turn; confirmed that `ApplyBranchDropTaperAcrossSegments(...)` multiplied by the drop factor segment-by-segment, which made runs shrink too aggressively through elbows/bends.
- New rule: branch drop applies at the connection between runs only. Once a run has inherited its root width, all unprotected segments in that same run hold that same width. Cascade-stop segments remain protected.
- Changed `ApplyBranchDropTaperAcrossSegments(...)` so it fills non-stopped segments with the inherited root width instead of repeatedly multiplying `currentWidth *= taperRatio`.
- Expected example with 15% drop: `unit feeder 12 -> middle trunk 10.2 -> register branch 8.67`, and a turn inside the register branch stays `8.67` instead of continuing to `7.37`.
- Built successfully with the existing `sourceArt` warning only. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 04:15:53`; installed plugin timestamp is `2026-05-20 04:15:20`.

## 2026-05-20 Unit Endpoint Normalizer Stale Tiny Width Fix

- User asked why the last modified ductwork was still tiny even though the general branch-drop behavior looked better.
- Latest log showed the affected run was `emory-1779256948269-41`. It had `startSegment=3`, `bodyWidth=12`, `cascadeStops=[1]`, and stale widths `[4.116,5.88,12,12]` from the old per-turn taper. The user then applied width `5.95` to one selected segment, so the rebuilt path logged `[4.116,5.95,12,12]`.
- Root cause: `NormalizeBlueOrangeUnitEndpointTaper(...)` was a separate remaining code path that still walked from the unit endpoint using repeated taper multiplication. It was also preserving/allowing stale values below the correct one-drop width.
- Changed `NormalizeBlueOrangeUnitEndpointTaper(...)` to compute the dropped width once from the unit/root segment and reuse that value for the rest of the run. It now also raises stale stored values below that one-drop width back to the correct one-drop width. This matches the new rule: taper at run/branch handoff, not at every turn.
- Build passed with the existing `sourceArt` warning only. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 04:24:31`; installed plugin timestamp is `2026-05-20 04:24:07`.

## 2026-05-20 Register-to-Register Blue Segment Width and Round Elbow Trim Fix

- User showed a visual case where the blue ductwork still looked like it tapered too aggressively. Latest log showed `sourceId=emory-1779256948269-41` at `[8.4,8.4,12,12]`, then register-ended same-blue segment `emory-1779256948269-40` inherited another drop from `8.4` to `5.88`.
- Root cause: same-layer segment-intersection inheritance treated a register-to-register blue crossing run as a new branch and multiplied by the branch-drop factor again. That made the first connected register-ended segment smaller than the line it connected to.
- Changed same-layer Blue/Orange segment intersections where the target run has registers at both endpoints to copy the connected segment width directly instead of applying another reduction. Log line `Emory intersection branch inherited...` now includes `reduced=0` for this copy-width case.
- Adjusted round elbow trim for mismatched segment widths: when a round connector bridges different widths, both sides now get at least the larger joint trim distance. This keeps the smaller-side curved elbow from tucking too tightly into the turn and exaggerating the taper visually.
- Built `EmoryDuctwork.aip` successfully with the existing `sourceArt` warning only. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 04:37:13`.

## 2026-05-20 Same-Color Segment Handoff Anchor Fix

- User changed the width of the problem line so it could be identified in the log. Latest log identified `sourceId=emory-1779256948269-42`: it has registers at both ends, intersects same-color source `emory-1779256948269-41`, and was being marker-synced from `41` at full `12` width with `targetSegment=0`.
- Root cause: the same-layer marker sync used the raw connection segment as the target start, even when that segment was an omitted terminal/register stub. For this case, segment `0` is omitted and the usable handoff/start should be the adjacent visible/control segment. The marker sync also copied the unit-connected parent width directly instead of applying the branch-drop factor to the register-to-register target.
- Added `ResolveInheritedConnectionAnchorSegmentIndex(...)` and routed same-layer segment inheritance plus same-layer marker sync through it. A connection landing on an omitted terminal segment now anchors the stored start/root width on the adjacent control segment.
- Added same-layer marker branch-drop handling for Blue/Orange unit-source to register-pair targets. New log line: `Emory same-layer marker unit branch drop source=<id> target=<id> rawRootWidth=<w> rootWidth=<w*factor> factor=<factor>`.
- Reverted the previous register-pair same-layer special case that copied width at 100%; same-layer segment intersections now apply the normal inherited branch drop again, but from the corrected handoff anchor.
- Built `EmoryDuctwork.aip` successfully with the existing `sourceArt` warning only. Deployed through the popup-preserving scheduled task; deploy log shows `Direct Emory deploy complete` at `2026-05-20 04:48:01`; installed plugin timestamp verified as `2026-05-20 04:47:28`.

## 2026-05-20 Unit Handoff Main Start Width Sync Fix

- User reported the Blue ductwork start at a unit was not matching the Green ductwork width at the same unit. Latest log showed Green source `emory-1779256948273-79` at the unit as `7.2`, while paired Blue source `emory-1779258110892-660` remained `12` because its unit endpoint segment was in `cascadeStops=[2]`.
- Changed `SyncEndpointOnlyWidthFromState(...)` to accept a narrow `ignoreCascadeStops` flag. Green/Light-Orange to Blue/Orange unit handoff sync now uses that flag so a stored/locked main start width cannot override the transition duct width at the unit endpoint.
- Widened hardcoded Blue-only unit handoff checks to `IsBlueOrOrangeRunLayer(...)`, so Light Orange -> Orange inherits the same exact start-width behavior as Green -> Blue.
- Updated logs/messages from "blue start" to "main start" for this path. New relevant logs include `Emory direct unit-pair main start sync... mainLayer=<Blue/Orange Ductwork>` and `Emory width-apply synced main start from transition...`.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 04:57:08`; installed plugin timestamp verified as `2026-05-20 04:56:39`.

## 2026-05-20 Unit Handoff Main Root Retaper Fix

- User reported the Blue ductwork did match the unit endpoint but then jumped back to full start size after the bend. Latest log after the previous deploy confirmed the pattern: `emory-1779258110892-660` became `widths=[12,12,5.6]` with `startSegment=2`, and the user's manually sized example `emory-1779258110891-645` showed `widths=[8,5.6]` with the same unit-side root segment.
- Root cause: the unit-pair sync was still endpoint-only. It copied the Green/Light-Orange width into the Blue/Orange segment touching the unit, but did not retaper/fill the rest of that same main run from that segment.
- Added `retaperFromTargetEndpoint` to `SyncEndpointOnlyWidthFromState(...)`. Unit handoff calls pass `true`, so the main segment touching the unit becomes the explicit start/root segment and `ApplyBranchDropTaperAcrossSegments(...)` fills downstream non-protected segments from that root width. A state like `[12,12,5.6]` should now become `[5.6,5.6,5.6]` unless a downstream segment is explicitly protected.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 05:03:34`; installed plugin timestamp verified as `2026-05-20 05:03:05`.

## 2026-05-20 Reset Stored Widths Button and Same-Layer Reverse Inheritance Fix

- User reported the segment-to-segment same-layer inheritance was still failing on a manually sized Blue line, then clarified the Reset Stored Widths button had never worked.
- Latest log showed `sourceId=emory-1779256948269-42` was manually width-applied, but same-layer marker sync still only allowed marker-owner `41` to drive unmarked `42`; selected seed protection blocked the reverse case. Added reverse same-layer marker sync so a manually touched unmarked Blue/Orange run can push its dominant visible width back into the marker-owned run and clear stale cascade-stop basis values for that sync.
- Reset Stored Widths panel click now uses a delegated button handler, the button is explicitly `type="button"`, and the click writes a panel-side debug line before calling `MDUX_cppResetSelectedEmoryStoredWidths(...)`.
- Native Reset Stored Widths now expands to connected runs, removes stored segment width metadata on affected sources, clears cascade-stop locks, retapers selected roots, reapplies inherited branch widths, persists cascade stops, and reports how many stored width records/stops were cleared.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 05:19:28`; installed plugin timestamp verified as `2026-05-20 05:19:03`.

## 2026-05-20 Green Bend No-Taper Rule

- User requested no taper on Green ductwork at ordinary bends; taper should remain only at internal anchors.
- Added Green/Light Green bend normalization in `GenerateEmoryForPath(...)`: for non-marker, non-collinear bend vertices, both adjacent segments are normalized to the larger adjacent width so ordinary elbows do not create a width step.
- Preserved the intended exception for same-Green intersections with internal anchors. Same-layer marker detection now also treats a nearby internal vertex on the other same-layer path (`12 pt`) as a marker boundary, so a Green-to-Green crossing with an internal anchor on either intersecting line does not get flattened as a normal bend.
- New path log flag: `greenBendWidthsNormalized=1` when the Green bend equalizer changes stored widths.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 05:28:52`; installed plugin timestamp verified as `2026-05-20 05:28:28`.

## 2026-05-20 Green Branch No-Drop Rule

- User reported width dropping was still happening on Green branches and asked to remove width dropping on Green ductwork branches.
- Added `ShouldApplyBranchDropBetweenStates(...)` so Green/Light Green parent-to-branch inheritance copies the parent width instead of multiplying by the branch drop percentage.
- Applied that rule to endpoint-to-segment branch inheritance, same-layer segment intersection inheritance, and branch-root refresh. Green/Light Green branch inheritance logs now show `reduced=0` when the drop is suppressed.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 05:35:31`; installed plugin timestamp verified as `2026-05-20 05:34:19`.

## 2026-05-20 Reset Stored Widths Green Normalization Fix

- User reported Reset Stored Widths still left Green ductwork with inconsistent widths; selecting all Green and manually setting one width fixed it, so the reset was preserving old branch/root widths.
- Added immediate native logging for reset attempts: `Emory reset stored widths invoked...`, plus `no selected source ids` if the command reaches native code without a usable selection.
- Reset now clears stored segment widths, stored source body widths, and cascade-stop metadata across all centerline/backup art for affected source IDs. This prevents Process Ductwork from recovering stale widths right after the reset.
- Added Green/Light Green same-layer expansion for reset, then normalized the affected Green/Light Green network to the max current green width. Reset logs now include `Emory green reset same-layer expansion...` and `Emory green reset normalized sources=<n> width=<w>`.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 05:56:31`; installed plugin timestamp verified as `2026-05-20 05:47:17`.

## 2026-05-20 Reset Stored Widths Button Selection Fallback

- User reported Reset Stored Widths still did nothing. Latest native log showed fresh Process Ductwork runs at `05:59`, but no `Emory reset stored widths invoked...` line, so the reset button did not reach native code.
- Root cause found in panel UI: `setEmoryControlsEnabled(false)` disabled the Reset Stored Widths button whenever no Emory selection was active. That made the user's workflow `Reset Stored Widths > Select Ductwork > Process Ductwork` a dead click.
- Changed the panel so Reset Stored Widths is always clickable in Emory mode.
- Changed native reset so if no selected Emory source IDs are found, it resets all Emory source width controls in the document instead of failing. The log will show `Emory reset stored widths no selected source ids; resetting all document Emory sources` and `resetAll=1`.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 06:01:59`; installed plugin timestamp verified as `2026-05-20 06:01:36`; deployed `panel.js` timestamp verified as `2026-05-20 06:00:44`.

## 2026-05-20 Reset Stored Widths Click Handler Hardening

- User clarified they had selected all ductwork before clicking Reset Stored Widths, so the prior explanation was incomplete.
- Checked logs again: neither native `Emory reset stored widths invoked...` nor panel-side `Reset Stored Widths clicked...` appeared, so the panel click handler itself was not firing or was guarded out.
- Removed the handler guard that skipped work when the target button was disabled, force-enables the button when binding, and added a panel log line `Reset Stored Widths handler fired disabled=false` before calling native reset.
- Panel-only deploy through popup-preserving scheduled task completed at `2026-05-20 06:04:58`; deployed `panel.js` timestamp verified as `2026-05-20 06:04:24`. Native plugin remained at timestamp `2026-05-20 06:01:36`.

## 2026-05-20 Branch Drop Selection Narrowing

- User reported changing branch drop percentage became extremely slow and was applying to all Blue ductwork, not just the selected branch.
- Latest log showed the problem clearly: branch-drop had `selectedSegments=5` but `selectedSources=189`, then inherited widths across `sources=153`. The exact selected generated segments were being collected, but the broader source collector was still used as the authoritative selection and pulled in parent/group linked source IDs.
- Changed `ApplySelectedEmoryBranchTaperReductionPercent(...)` so selected generated Emory segment source IDs are authoritative. The broad `CollectSelectedEmorySourceIds(...)` result is now only a fallback when no generated ductwork segment is selected, such as centerline-only selection.
- Added branch-drop log fields `rawSelectedSources=<n>` and `generatedSeedSources=<n>` so future logs show whether the command is staying scoped. For a selected generated branch, `selectedSources` should now match `generatedSeedSources`, not the broad raw count.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 06:22:32`; installed plugin timestamp verified as `2026-05-20 06:18:28`.

## 2026-05-20 Segment Width Lock Buttons

- User asked for explicit Lock Segment and Unlock Segment buttons so one selected segment can keep its current width instead of being overwritten by cascading/taper inheritance.
- Reused the existing native protected-segment/cascade-stop mechanism because inheritance already checks it and leaves marked segments unchanged. No second metadata system was added.
- Moved the segment protection controls into the main Emory slider area under Branch Drop, labeled `Lock Segment Width` and `Unlock Segment`. Removed the duplicate stop-marker row from the collapsed Emory Tools area.
- Updated panel status text from `Stop` / `Stop marked` to `Locked` / `Locked segment`, and changed native messages so the button reports `Locked segment... Cascading width changes will leave this segment alone.`
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 06:51:20`; installed plugin timestamp verified as `2026-05-20 06:50:12`.

## 2026-05-20 Unlock Old Stop Segment Compatibility Fix

- User reported previously marked stop segments did not unlock with the new `Unlock Segment` button.
- Latest log still showed the test Green run carrying `cascadeStops=[1]`, but no native unlock action was logged afterward, so the panel was likely disabling or failing to recognize the old mark.
- Changed selection-state detection to look for locked/stop metadata on both the visible source centerline and its backup centerline mapping. This makes old marks and backup-mapped marks show as unlockable.
- Changed native unlock so it removes the selected segment from both canonical/backup metadata and visible source metadata. It no longer fails if only one side has the old mark.
- Changed the panel so `Unlock Segment` is enabled whenever exactly one generated Emory segment is selected, even if old metadata detection is imperfect. Native code now decides whether anything was actually locked.
- Added native log line `Emory segment width lock action enabled=<0/1> ... canonicalMarkedBefore=<0/1> visibleMarkedBefore=<0/1>` for future debugging.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 06:59:36`; installed plugin timestamp verified as `2026-05-20 06:58:55`; deployed `panel.js` timestamp verified as `2026-05-20 06:58:39`.

## 2026-05-20 Segment Lock Button Click Binding Hardening

- User reported unlock still did not appear to work. Latest native log after the `06:59` deploy had no `Emory segment width lock action...` line at all, only repeated width-apply logs, so the button click was not reaching native code.
- Replaced direct click bindings for `Lock Segment Width` / `Unlock Segment` with one delegated document click handler, matching the more reliable Reset Stored Widths binding pattern.
- The delegated handler force-enables the clicked button, logs `Segment width lock handler fired action=lock|unlock`, then calls the same native bridge methods.
- Panel-only deploy through the popup-preserving scheduled task completed at `2026-05-20 07:09:32`; deployed `panel.js` timestamp verified as `2026-05-20 07:08:51`. Native plugin remained at timestamp `2026-05-20 06:58:55`.

## 2026-05-20 Segment Lock Buttons Kept Clickable

- User reported unlock still did not work. Latest logs after the `07:09` deploy still had no panel `Segment width lock handler fired...` line and no native `Emory segment width lock action...` line, only width apply calls.
- Root cause: disabled buttons do not emit click events, so the delegated handler could not catch the click if selection-state refresh disabled `Unlock Segment`.
- Changed `setEmoryControlsEnabled(...)` and selection-state refresh so `Lock Segment Width` and `Unlock Segment` remain clickable. Native code will validate the selection and report if nothing is selected/locked.
- Panel-only deploy through the popup-preserving scheduled task completed at `2026-05-20 07:12:43`; deployed `panel.js` timestamp verified as `2026-05-20 07:11:52`. Native plugin remained at timestamp `2026-05-20 06:58:55`.

## 2026-05-20 T Junction Near Elbow Clamp Fix

- User reported T junctions were too close to elbows, with connector ends overlapping or extending past each other instead of clamping cleanly.
- Root cause: network connector clamp anchors only considered other network T/cross connection centers. Ordinary internal elbow/corner trim zones were invisible to the T connector arm builder, so a nearby T arm could run into the elbow connector region.
- Added reserved-length elbow clamp anchors for internal non-collinear joints. T/cross connector arms now clamp before the elbow's reserved trim zone instead of using their full arm length when an elbow is nearby on the same segment.
- Existing T-to-T clamp behavior is preserved; only elbow anchors carry a reserved trim length. Normal neighboring connection anchors still use the midpoint clamp.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 07:29:40`; installed plugin timestamp verified as `2026-05-20 07:27:57`.

## 2026-05-20 Reset Stored Widths Selection Scope Fix

- User selected two generated duct lines and clicked Reset Stored Widths, but it altered the whole document.
- Latest native log showed the bug clearly: `selectedSegments=3` but `selectedSources=259`, so the selected generated segments were detected, then the reset used the broad parent/group source collector anyway.
- Changed `ResetSelectedEmoryStoredWidths(...)` to mirror the branch-drop selection fix: generated Emory segment source IDs are authoritative when present; the broad selected-source collector is only a fallback for source centerline selection.
- Reset logs now include `rawSelectedSources=<n>` and `generatedSeedSources=<n>`. For the same kind of generated segment selection, expected shape is `rawSelectedSources=259 generatedSeedSources=3 selectedSources=3` instead of touching all 259.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 07:53:53`; installed plugin timestamp verified as `2026-05-20 07:53:16`.

## 2026-05-20 Green Manual Width Downsize Fix

- User could not reduce the selected generated duct width from an oversized Green run. Latest log showed width-apply accepted `newWidth=2`, but the regenerated Green path still logged `widths=[45,45,45]` with `greenBendWidthsNormalized=1`.
- Root cause: the Green bend no-taper rule normalized ordinary Green bends back to the larger adjacent width after manual width apply. That prevented accidental Green bend taper, but also blocked intentional manual downsizing on a selected Green run.
- Changed manual width application for Green/Light Green states so setting a generated segment width applies that width across the full Green source/run. This keeps the no-taper-at-Green-bends rule intact while allowing the selected Green run to actually get smaller.
- Future note: we may add tapering back to Green ductwork on curves later, but it needs to be explicit/controlled so ordinary Green bends don't accidentally shrink and manual width changes don't get overwritten.
- Build passed with the existing `sourceArt` warning only. Deploy through the popup-preserving scheduled task completed at `2026-05-20 07:57:40`; installed plugin timestamp verified as `2026-05-20 07:57:18`.
