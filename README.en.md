🌐 Language / 语言：[简体中文](README.md) ｜ **English**

# CST-LinXi-Macro

A **slow-wave structure user-defined watch** macro for CST Studio Suite eigenmode
simulations. It attaches to a `phase` parameter sweep, extracts the eigenmode field
data at every phase point, computes the **normalized phase velocity** and the
**Pierce interaction impedance** of a periodic slow-wave structure, and produces the
**Brillouin dispersion**, **normalized phase velocity** and **coupling impedance**
curves automatically.

- License: [MIT](LICENSE)
- Supported platforms: CST Studio Suite 2024 / 2025 / 2026 (64-bit)
- Current version: `v7.9.5.4 (Update 4)`

> I am the author of this project. What follows is both the user guide and a record of the
> design decisions behind it: steps that you carry out are written as "you", while project
> decisions and copyright are written as "I".

> ⚠️ **The three things that bite most often** (each one is explained in its section below):
> ① `LinXi.ini` must stay **ANSI/GBK** encoded — saving it as UTF-8 fails immediately;
> ② installation and project paths **must not contain parentheses**, and shorter is better;
> ③ **do not touch the CST window during a sweep**, and **clear old results before re-sweeping**.

---

## 1. Repository layout

```
CST-LinXi-Macro/
├─ zh-CN v7.9.5 Update 4/          # Chinese build
│  ├─ 双击我自动安装.bat             # Chinese one-click installer
│  ├─ 双击我自动卸载.bat             # Chinese one-click uninstaller
│  └─ MacroKit/
│     ├─ LinXi.bas                 # main macro: the sweep watch (core computation flow)
│     ├─ Me.bas                    # setup wizard: adds the watch to a project
│     ├─ She.bas                   # version check macro
│     ├─ LinXi.ini                 # version info (must be ANSI/GBK encoded)
│     └─ LinXi.dll                 # core computation library (64-bit)
├─ en-US v7.9.5 Update 4/          # English build (same code generation)
│  ├─ Double‑click me for automatic installation.bat
│  ├─ Double‑click me for automatic uninstallation.bat
│  └─ MacroKit/                    # identical file set to the Chinese build
├─ LICENSE
├─ NOTICE
├─ README.md
├─ README.en.md
├─ .gitignore
└─ .gitattributes
```

> **Both directories hold the same code generation; only the language differs.** Their
> `LinXi.bas` / `Me.bas` / `She.bas` are structurally and logically identical: only the UI
> strings and log messages are in a different language (the Chinese build also carries full
> Chinese comments, the English build has none). Each build takes its version information from
> its own `LinXi.ini`, and both ship the same `LinXi.dll`.

---

## 2. Installation

**Double-click `en-US v7.9.5 Update 4\Double‑click me for automatic installation.bat`**
(the `‑` in the file name is U+2011, a non-breaking hyphen), or the Chinese installer if you
prefer the Chinese UI. Both deploy the same file set, and no manual copying is needed.

The installer will:

1. Request administrator rights (UAC; it continues automatically after elevation).
2. Detect installed CST versions from the registry and accept **2024 or later** only;
   when several versions are present, MacroKit is installed into **all of them without asking**.
3. Deploy the files as follows:

| Source | Destination | Deployed name |
| --- | --- | --- |
| `MacroKit\LinXi.dll` | `<CST>\AMD64\` | `LinXi.dll` |
| `MacroKit\LinXi.ini` | `<CST>\AMD64\` | `LinXi.ini` |
| `MacroKit\LinXi.bas` | `<CST>\AMD64\` | `LinXi_Watch.bas` ← **the watch macro itself** |
| `MacroKit\LinXi.bas` | `…\Library\Macros\Admin LinXi Macro\` | `Run LinXi Macro.mcr` |
| `MacroKit\Me.bas` | same as above | `Define LinXi Macro.mcr` |
| `MacroKit\She.bas` | same as above | `Check LinXi Version.mcr` |

Here `…\Library\Macros\` means
`%APPDATA%\Dassault Systemes\CST STUDIO SUITE\Library\Macros\`.

The three files under `AMD64` are deployed **once per detected CST version** (the script prints
`[ k/n ] CST installation root: …` for each of them). The user macro directory does not depend on
the CST version, so it is deployed only once.

**Key design point:** the watch macro `LinXi_Watch.bas`, the core library `LinXi.dll` and
the config file `LinXi.ini` all live in the same `AMD64` directory. The per-project
`Model.pfc` is copied **on demand** by the `Define LinXi Macro` macro when the watch is
added; the installer never writes into a project directory.

To uninstall, run the uninstaller. It removes the three files under `AMD64` of every detected
CST 2024+ version and **deletes the whole `Admin LinXi Macro` directory**, so there is nothing to
clean up file by file.

> The uninstaller **never touches your projects**.

> ⚠️ **The path must not contain parentheses.** Keep your projects in short paths too (Windows 260-character limit).

---

## 3. Usage

1. Start CST and open an eigenmode project (periodic boundary set on the slow-wave structure).
2. Add `Define LinXi Macro` to the project macros (from the `Admin LinXi Macro` menu).
3. The setup wizard asks for four things:
   - **Step 1** — modes to compute: all / ticked / custom list such as `1,3,5`;
   - **Step 2** — the coupling-impedance reference point (the periodic direction is
     locked to the centre of the calculation domain);
   - **Step 3** — power-flow method: **indirect** (from the E and H fields) or
     **native** (CST's own power-flow field);
   - **Step 4** — the **region** to compute: **forward-wave** (the frequency must rise with
     `phase`), **backward-wave** (the frequency must fall with `phase`), or **both**
     (no frequency-direction check, every sweep point is used).
4. Confirm, then start the `phase` parameter sweep — **no code changes are required**.
5. Results appear under `1D Results\PlotOutput LinXi Macro Result\`.
   Run the `Check LinXi Version` macro at any time to inspect the versions in use.

> Whichever region is selected, two kinds of invalid sweep points never take part in the
> calculation: multiples of π, and points whose frequency stayed exactly the same for
> 5 consecutive samples (which means the phase difference never reached the periodic boundary).

**Installation success criterion:** the CST message window shows
`核心计算库 LinXi.dll 加载成功 v122` on the first run (the number follows the actual DLL version).

### Parameters the macro writes into the project

| Parameter | Meaning |
| --- | --- |
| `Macro_SweepWatch_Enable` | Enable flag, written as `integer part.decimal digit`: the integer part selects the power-flow method, the decimal digit selects the region |
| ↳ integer part | `1` = indirect power flow, `2` = native power flow, `0` = disabled |
| ↳ decimal digit | `0` = forward- and backward-wave regions, `1` = forward-wave region only, `2` = backward-wave region only |
| ↳ valid values | `1.0` `1.1` `1.2` `2.0` `2.1` `2.2`; any other value (such as `1.3`, `1.11`, `3.0`) makes the wizard appear again on the next run |
| `Kc_RefPos_x` / `y` / `z` | Reference-point coordinates (periodic direction is the domain centre) |
| `Macro_Modex` | Interaction-impedance flag for mode `x`. **Only its presence matters, not its value** |

`1.0` / `2.0` mean "compute both regions". An integer `1` / `2` left behind by Update 2 is
interpreted as `1.0` / `2.0` as well, so an upgraded project computes the backward-wave region
too; delete `Macro_SweepWatch_Enable` to let the wizard appear again and pick `1.1` / `2.1`
when only the forward-wave region is wanted.

Delete `Macro_SweepWatch_Enable` (or `Kc_RefPos_*`, `Macro_Mode*`) to make the wizard
appear again on the next run.

> These parameters **must not be swept**; doing so corrupts the macro's state detection.

---

## 4. Precautions during a sweep

The macro reads and writes CST UI data continuously while running, so any manual
interaction can "fight it for control" and break the field reading.

- **Do not operate the CST window during a sweep.**
- Use a **separate CST window** for the eigenmode sweep.
- **Never run several eigenmode sweeps in the same window at once.**
- Make sure the electric-field results are ready before sweeping.
- Do not open the `Fields on Plane` or `Cutting Plane` views of E/H in eigenmode projects.
- **Clear existing results before repeating a sweep**, otherwise the previous results may be lost.
- The hexahedral **JDM** solver is recommended (more accurate); under JDM, eigenmode
  names must not contain a decimal point.

---

## 5. Reading the results

| Curve | Result tree location | Notes |
| --- | --- | --- |
| Brillouin diagram (β–f) | `Brillouin Diagram Beta` | Should be smooth and continuous, with no jumps or breaks |
| Brillouin diagram (phase–f) | `Brillouin Diagram Phase` | Sweep phase versus frequency |
| Normalized phase velocity vp/c | `Normalized Phase Velocity` | Should follow the expected trend with frequency |
| Pierce interaction impedance Kc | `Pierce Interaction Impedance` | Usually has fewer points than the β curve: invalid points and points outside the selected region are skipped, which is normal |

Curves with cliffs, spikes or discontinuities caused by several consecutive `0` values are called
**malformed curves**; they are almost always caused by a failed field reading.

> With "forward-wave region only" or "backward-wave region only", only the sweep points whose
> frequency moves in the required direction appear on the Kc curve, so it naturally holds fewer
> points than the Brillouin curve. With "both regions", every sweep point except the invalid
> ones takes part in the calculation.

---

## 6. Troubleshooting

The full run log is written to `Temp\_macro_log.txt` **in the project directory** (rotated to
`*.bak` above 5 MB). The CST message window shows a summary prefixed with
`MacroMsg [INFO] / [WARN] / [ERROR] / [CRIT]`.

| Symptom | Cause and fix |
| --- | --- |
| Config file not found / 8 keys missing | `LinXi.ini` is missing or was saved as **UTF-8**. It must be **ANSI/GBK**; re-run the installer to restore it |
| DLL fails to load / too old | `LinXi.dll` not deployed, replaced by a 32-bit build, or older than `DllMinVersion` in `LinXi.ini`. Re-install |
| Malformed curves | Usually a failed field reading: the UI was touched during the sweep, several tasks shared one window, or the delay does not match the machine speed. Increase `WAIT_SEC` in `LinXi.bas`, or re-sweep the affected range |
| Log reports "frequency did not change with phase" and invalidates the sweep | The solver is not applying `phase` to the periodic-boundary phase shift. Check: ① both boundaries of that direction are `periodic`; ② the phase shift is bound to the sweep variable `phase`; ③ the domain is exactly one period long |
| Log reports "frequency did not rise" | The selected region is **forward-wave only** and this sweep point's frequency did not rise with `phase`, so its interaction impedance is skipped by design (frequency and phase are still recorded). If such points should be computed, choose the backward-wave region or "both regions" in the wizard |
| Log reports "frequency did not fall" | The selected region is **backward-wave only** and this sweep point's frequency did not fall with `phase`, so it is skipped as well. Fix it the same way as above |
| The setup wizard appears on every run | `Macro_SweepWatch_Enable` is not one of `1.0` `1.1` `1.2` `2.0` `2.1` `2.2` (for example it was edited to `1.3` or `1.11`), or the parameter was deleted. Simply choose again in the wizard; the macro writes a valid value back into the parameter list |
| Temporary-file write failures / path too long | The project path is too deep for the Windows 260-character limit; move it somewhere shorter |

---

## 7. Development

> **If you want to build on this project, start from `v7.9.5.4` (i.e. Update 4) or later.**

- The sources are CST embedded VBA (`.bas`) encoded as **ANSI/GBK** with **CRLF** line
  endings. Keep both properties when editing, or CST will mis-read Chinese text and key values.
- The Chinese `.bas` sources carry full Chinese comments, written with **trailing single quotes**.
- **The English `.bas` sources carry no comments**, apart from the MIT copyright block at the
  top of each file (the MIT licence requires the copyright and permission notice to be kept).
- In both builds, `LinXi.bas` groups the tunable constants
  (`WAIT_SEC`, `MAX_FREQ_RETRY`, `FREQ_FLAT_*`, `KC_SANITY_MAX`, `MAX_PATH_SAFE`, …) at the top
  of the file so they can be adapted to a given machine.
- The encoding of `Macro_SweepWatch_Enable` lives in the `PW_SRC_*` / `REGION_*` constants and the
  `DecodeEnableFlag` function at the top of the file: the integer part is the power-flow method and
  the single decimal digit is the region. Adding a region means extending that encoding plus `RegionText`.
- The two builds must stay logically identical: whenever `LinXi.bas` changes, update the Chinese and
  the English version together (the English one carries no comments).
- `LinXi.dll` is likewise my own work and is released under the same MIT licence, but I
  ship it **as a prebuilt binary only, without build sources**, which means the DLL's internal
  logic cannot be modified or recompiled by anyone but me. Its exported prototypes are declared
  as `Declare Function` blocks at the top of `LinXi.bas`. See [NOTICE](NOTICE).
- **This repository does not contain the sample macros `LinXi1.bas` / `Me1.bas`** that ship with
  CST: their copyright belongs to Dassault Systèmes and they are outside this project's MIT grant.
  See [NOTICE](NOTICE).
- There is no automated test suite yet. After any change, run a full `phase` sweep on a
  small model in CST and check both the log and the resulting curves.

Issues and pull requests are welcome (Chinese templates live under `.gitee/`).

---

## 8. Changelog

### Update 4 — `v7.9.5.4` (2026-09-24)

Only `LinXi.bas` changed in this release (`Me.bas`, `She.bas`, `LinXi.dll` and
`LinXi.ini` are untouched). All of it is robustness and usability work; the sweep
flow and the result algorithms are unchanged.

- **Unified the LICENSE attribution:** its copyright notice is no longer inconsistent.
- **A too-deep project path now aborts before anything is written**, naming the
  `Temp` path length and the characters still available for the file tag, and
  telling you to move the project to a shorter directory. Previously an over-long
  path could write files that the cleanup pass could no longer remove.
- **Fixed the illegal-character substitution for parameter-set keys**, which had
  never taken effect: the search strings did not match the separators actually
  used when building the key (one space short). Projects whose parameters contain
  e.g. a decimal point now get a correctly rewritten file name instead of relying
  on the tag happening to fit.
- `Dir()` is now used only for the wildcard batch delete. Existence checks use
  `GetAttr` instead, so they cannot disturb `Dir`'s enumeration state and leave
  part of a batch uncleaned.
- In the mode-selection group, **clicking a check box no longer throws focus back
  to Mode 1** (the group is only switched when it is not already "selected modes
  only"), so consecutive keyboard selection is no longer interrupted.
- Added the internal constants `PW_SRC_INDIRECT` / `PW_SRC_NATIVE`, replacing the
  scattered integer comparisons. The old same-named constants meant the opposite
  of the internal numbering, which was a trap for future edits.
- **Split two warning flags apart.** "Frequency does not change with `phase`" and
  "`Kc` exceeds the sanity limit" shared one flag, so once the first had been
  reported the second could never appear. They are independent now.
- **Replaced the 25 list bullets `·` in the Help dialog with an ASCII `-`.**
  `·` (U+00B7) renders under both GBK and CP1252, but keeping the English build
  pure ASCII is safer.
- The English `LinXi.bas` mirrors all of the above.
- Both READMEs were updated to this version.

### Update 3 — `v7.9.5.3` (2026-09-23)

- The setup wizard gained a **Step 4**: choose the **region** whose interaction impedance is computed.
  - **forward- and backward-wave regions** — no frequency-direction check, every sweep point is used;
  - **forward-wave region only** — only points whose frequency rises with `phase` are used;
  - **backward-wave region only** — only points whose frequency falls with `phase` are used;
  - only the description of the **currently selected** option is shown; the other two are hidden.
- `Macro_SweepWatch_Enable` grew from an integer into an **integer part plus one decimal digit**:
  the integer part selects the power-flow method, the decimal digit selects the region. An invalid
  value (such as `1.3` or `1.11`) makes the wizard appear again on the next run instead of aborting.
  An integer `1` / `2` from an older project is read as `1.0` / `2.0`, i.e. "both regions".
- New and updated log messages: the final summary prints the power-flow method and the region, and
  points rejected by the direction check are reported individually.
- The **Help** dialog now uses a larger vertical line spacing (22 → 28); the wording is unchanged and
  the dialog grew from 540 to 645 units in height.
- The installer no longer asks which CST version to use: every detected version from 2024 on is
  installed into, and the shared user macro directory is deployed only once.
- The English build mirrors all of the above (and stays comment-free); both READMEs were updated.

---

## 9. License

Released under the **MIT License** — see [LICENSE](LICENSE).

```
Copyright (c) 2026 Limorazp
```

You may use, modify and redistribute this software, **including commercially**, as long
as the original copyright and permission notices are retained. The software is provided
"as is", with no guarantee of accuracy; validate key simulation results yourself.

> `LinXi.dll` is distributed as a prebuilt binary that I compiled myself
> and links no third-party computational library. See [NOTICE](NOTICE) for details.
