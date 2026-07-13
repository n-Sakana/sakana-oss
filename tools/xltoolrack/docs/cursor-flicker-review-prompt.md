# External review prompt: Excel FE cursor flicker

You are taking over an unresolved Excel/VBA architecture problem in:

`C:\repos\pub\xltoolrack`

Read these files first, completely:

1. `docs/edit-freeze-and-fe-pump.md`
2. `docs/HANDOFF-codex.md`
3. `src/addin/JobPump.bas`
4. `src/common/JobHost.cls`
5. `src/common/AppStateSnapshot.cls`
6. `src/common/Infra_Dispatch.bas`
7. `src/tools/stopwatch.bas`
8. `src/common/ChannelFile.bas`

Do not edit, build, install, kill Excel, or force any window to the foreground until you have written an evidence-based diagnosis and an experiment plan. The main worktree is intentionally dirty with another agent's and the user's work. Never reset, clean, or discard it.

## Goal

Keep the cell-edit freeze permanently fixed while eliminating, or reducing below perception, the intermittent Windows "working in background" cursor (`IDC_APPSTARTING`, arrow plus spinner) shown while timer tools run.

Excel must retain all of its own contextual pointers:

- white thick cross over ordinary cells;
- black thin cross over the fill handle;
- I-beam while editing.

## Hard constraints

- Shipped VBA must not use Win32 `Declare`, Shell, WMI, `taskkill`, or an external helper process.
- A diagnostic observer may use Win32 outside the shipped add-in, but that is not a product solution.
- A BE Excel worker must never synchronously call the FE object model or write an FE `Range` through COM.
- Cell editing must never freeze, and ESC/input must remain available.
- Do not force `xlNorthwestArrow`, `xlWait`, or another synthetic pointer over Excel's contextual pointers.
- Do not infer success from a background/hidden Excel automation test. The cursor symptom exists when Excel is foreground.
- Change one variable per experiment and retain a known-good rollback build.

## Current architecture

There are separate hidden BE Excel processes. Each worker computes once per second and publishes its latest result to one session aggregate file under `%TEMP%`.

The user-facing FE Excel owns a one-second `Application.OnTime` chain:

```text
JobPump.Pump_Tick
  DoEvents
  HostServices.Jobs().PumpOnce
    Application.EnableEvents = False
    ChannelFile.ReadAggregate                 (one file open for all jobs)
    for each fresh record
      update JobState memory cache
      Infra_DispatchResult
        tool.OnResult
          stopwatch: Range(...).Value2 = displayValues
    exceptional error checks / stale-or-terminal sweep
    restore application state except cursor
  schedule next OnTime
  DoEvents
```

Excel defers `OnTime` while it cannot run VBA, including cell-edit mode in observed use. Thus FE-local `Value2` writes happen only after editing ends. This direction reversal fixed the original permanent edit freeze.

## Confirmed facts

1. The old BE-to-FE synchronous COM write architecture could park/retry against the FE STA during editing and starve keyboard/ESC. It is not an acceptable direction.
2. The currently registered desktop add-in is the intended file: `C:\repos\pub\xltoolrack\dist\xltoolrack.xlam`.
3. Current shipped `src` contains no `Application.Cursor` reference. `AppStateSnapshot` deliberately excludes cursor state, and the build validator rejects future cursor overrides.
4. Removing cursor save/restore fixed a separate real bug: an isolated old `AppStateSnapshot` could capture a transient busy cursor late and keep it busy. It did **not** eliminate the user's ordinary intermittent flicker on desktop.
5. The user still sees obvious flicker with the correct build loaded, approximately once every 2-3 seconds during Multi Stopwatch.
6. Removing both `DoEvents` calls from `Pump_Tick` made the desktop symptom **clearly worse**. That experiment has been rolled back. Do not repeat it as a proposed fix without new evidence.
7. Pinning `xlNorthwestArrow` reduced sampled spinner onsets but broke the native white cross/fill-handle behavior and is rejected.
8. Aggregate file reads, memory caching, shared application-state snapshots, and removal of steady-state sweeps reduced callback work, but did not remove the user's residual symptom.

## Important failed/misleading experiments

### Cursor snapshot conclusion was overgeneralized

On note, a test-only full `AppStateSnapshot` produced a 100% busy cursor, while the same guard without cursor state produced 0. Removing cursor state was correct, but the agent wrongly concluded that this explained every flicker. Desktop use disproved that conclusion.

### Note foreground sampling was not sufficient proof

Post-fix note sampling reported 0 onsets for 30-second normal/guard/normal windows. Desktop still visibly flickered. Treat device cursor scheme, foreground behavior, sampling rate, timer cadence, and measurement interference as uncontrolled variables. User observation on the affected desktop is the positive case.

### Removing `DoEvents` was a regression

The hypothesis was that `DoEvents` allowed Excel to paint its transient busy state. In actual desktop use, removing both yields made flicker clearly worse. The prior binary was restored.

Targeted note regressions for the no-`DoEvents` variant passed cursor contract, result dispatch, and a 22-second FE-busy test. The E2E core reached seven jobs and kept 30 cell inputs responsive, but note exhausted resources while adding jobs 9/10 and Helm went offline. This is not a completed all-green run and is not evidence that the variant was good.

## Current state

- Desktop installation: rolled back to the pre-atomic-callback build.
- Installed host SHA-256: `0AE109F5F7BE0D904AE6DCAA424769CA7E6843B86A7EB864D4675C0199909746`.
- Desktop Excel process count after rollback: 0.
- Source `JobPump.bas`: both `DoEvents` calls restored.
- Temporary review branch: `codex/cursor-probe-20260712-0952`, current tip `8b0614e4161155e093b6ee63f70f347a5aab7ea9`.
- `main` remains intentionally uncommitted/dirty; do not use the temporary branch history as a clean final commit plan.

## What your review must answer

1. Give the exact callback timeline and identify which operations can cause Excel/Windows to select `IDC_APPSTARTING`, distinguishing documented fact, local measurement, and inference.
2. Explain what `Application.Cursor` can and cannot represent. In particular, determine whether it can ever snapshot the actual white cross/black fill-handle/I-beam shape. Do not assume it can.
3. Design a foreground-positive experiment that isolates, one at a time:
   - `OnTime` entry/return only;
   - the two `DoEvents` calls separately;
   - `Application.EnableEvents` toggling;
   - aggregate open/decode;
   - `AppStateSnapshot` property reads/restores (without cursor);
   - `Infra_DispatchResult` setup/status/log work;
   - `Range.Value2` assignment;
   - recalculation/screen invalidation after `Value2`;
   - next-`OnTime` scheduling and callback return.
4. The experiment must not poll the live FE through COM during the sampling interval, because that changes the system being measured. Prefer a test-only stage selector set before sampling, or an external file-controlled selector, with no steady-state diagnostic I/O.
5. Propose the smallest safe product changes in priority order. Include an explicit rollback criterion for each.
6. Challenge the assumption that every worker result must be visibly written once per second. Consider coalescing, update-on-user-idle, calculation/manual repaint effects, a single batched `Value2` write, shape/status-bar display, or other pure Office-object-model approaches, but preserve correct semantics and explain tradeoffs.
7. State whether zero spinner is logically possible under the stated constraints, and what evidence would be required before claiming a hard limit.

Return a concise diagnosis, an experiment matrix, and a recommended first experiment. Do not implement until the user approves the plan.
