# Cursor flicker fix result (2026-07-12)

## Outcome

The original acceptance condition was **not achieved**:

- required: desktop, three tools, WAIT duration P95 < 10ms
- required: owner confirms the flicker is no longer visible
- actual: every accepted 90-second candidate remained at tick P95 >= 15.625ms
- owner result: no improvement; the full coalesce/profile run felt clearly worse

The failed runtime changes were rolled back. The installed candidate retains the test-only T0 profiler and disables
T1 flush dispatch at runtime.

## Desktop A/B (profile mode, three tools / seven jobs)

| Runtime | Window / ticks | Tick P50 | Tick P95 | Tick max | Aggregate P95 | Non-read P95 | WAIT onsets | WAIT P50 | WAIT P95 |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| stopwatch coalesce; pi row writes | 90s / 77 | 7.812ms | 15.625ms | 19.531ms | 7.812ms | 7.812ms | 3 | 12.6ms | 16.9ms |
| stopwatch + pi coalesce | 90s / 24 | 15.625ms | 19.531ms | 23.438ms | 7.812ms | 11.719ms | 12 | 17.0ms | 32.5ms |
| no coalesce (rollback) | 90s / 39 | 11.719ms | 15.625ms | 15.625ms | 7.812ms | 11.719ms | 2 | 21.9ms | 30.1ms |

Raw evidence:

- `2026-07-12-before-pi-coalesce-foreground/`
- `2026-07-12-after-dynamic-buffer-controlled/`
- `2026-07-12-t1-rollback-controlled/`

Tick counts differ because Excel intentionally defers `Application.OnTime` while it is not ready. Percentiles above
are over callbacks that actually executed. All three windows kept the same visible FE in the foreground and ended with
seven active jobs. Profile mode adds instrumentation overhead, so these values are for relative A/B decisions, not a
claim that the normal distribution has the same absolute cost.

## Follow-up one-variable trials

| Trial | Tick P95 | Aggregate P95 | Dispatch P95 | Decision |
|---|---:|---:|---:|---|
| rollback stage baseline | 15.625ms | 7.812ms | 7.812ms | reference |
| lazy Logger/Status/context services | 19.531ms | 11.719ms | 7.812ms | rolled back |
| channel files in `%TEMP%\xltoolrack\` | 15.625ms | 7.812ms | 7.812ms | no improvement; rolled back |
| compact life payload (420 to 21 fields) | 19.531ms | 7.812ms | 7.812ms | rolled back |

Raw evidence:

- `2026-07-12-stage-profile/`
- `2026-07-12-stage-profile-lazy-context/`
- `2026-07-12-stage-profile-channel-subdir/`
- `2026-07-12-stage-profile-life-packed/`

Error polling and cleanup both measured P95 0ms in the stage baseline. Aggregate decode/read and result dispatch were
the only measured nonzero PumpOnce stages.

## Build and automated suite

- installed xlam SHA-256: `F09F5916A5059B5A396C213EBDFC74D75A172520B2960C5C4EA04A134C01D02F`
- installed worker SHA-256: `BBBEB7369696EE4E59C98FD899332A84B9220A998EF7F7C22450CD876F099D8A`
- 15-case `test/Run-All.ps1`: exit code 0, `ALL COM HARNESSES PASSED`
- complete output: `2026-07-12-full-suite.log`

Environment caveat: after all `test-ribbon-ui` assertions passed, closing its temporary xlam displayed the known VBA
`File not found` modal. The modal was dismissed externally and only that test-owned Excel PID was stopped; the suite
then continued and returned 0. Thus all assertions are green, but that case is not unattended-clean on this machine.

## HostMain note

The current `HostMain.bas` diff is not a cursor-coalesce change. It provides the FE lease file, shutdown tombstone,
stale-session cleanup, and `JobPump.Pump_Stop` during shutdown so workers distinguish an editing/busy FE from a dead
FE. Temporary HostMain diagnostics used while examining the close modal were restored; no diagnostic-only branch
remains.

## Next authorized decision

The directive's remaining options are T4 and require owner approval:

1. change the FE tick from 1s to 2s (lower flicker opportunity, but stopwatch seconds can skip), or
2. pin a normal arrow during the callback (replaces the WAIT flash and requires a validator policy change).

Neither option was applied.
