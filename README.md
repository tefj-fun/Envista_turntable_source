# Envista Turntable Source

This repository contains the full **C# WinForms** solution for the Envista turntable inspection demo.  
It is the canonical source-of-truth for engineering changes; the matching production binaries are
published separately in [`Envista_turntable_stable`](https://github.com/tefj-fun/Envista_turntable_stable).

---

## Table of Contents

1. [Architecture Overview](#architecture-overview)  
2. [Repository Layout](#repository-layout)  
3. [Prerequisites](#prerequisites)  
4. [Building & Running](#building--running)  
5. [Hardware Setup](#hardware-setup)  
6. [Workflow Walkthrough](#workflow-walkthrough)  
7. [Logic Builder](#logic-builder)  
8. [Turntable & Cameras](#turntable--cameras)  
9. [Coding Guidelines](#coding-guidelines)  
10. [AI Assistant Usage](#ai-assistant-usage)  
11. [Release Process](#release-process)  
12. [Troubleshooting](#troubleshooting)

---

## Architecture Overview

The application orchestrates:

- **SolVision attachment model** (top camera) to locate attachment points.
- **SolVision defect model** (front camera) to evaluate each attachment.
- **ComXim turntable** for rotational positioning and homing.
- **Huaray cameras** (USB top + GigE front) for image acquisition.
- A WinForms UI with two major tabs: **Initialize** (hardware/model setup) and **Workflow** (execution, review, logic builder).

Key libraries:

- `Solvision.dll` / SolVision SDK (ExecuteType.Dll workflow).  
- `MVSDK_Net.dll` (Huaray camera SDK).  
- `Emgu.CV` for image manipulation/overlays.  
- `CefSharp` for embedded web components (logic builder docs/help popups).
- Custom `TurntableController` for COM-based serial commands.

---

## Repository Layout

```
DemoApp.sln
DemoApp/
├─ DemoApp.csproj
├─ DemoApp.cs / Designer / resx         → main WinForms Form
├─ Program.cs                            → entry point
├─ Properties/                           → Assembly info, resources, settings
├─ Dependencies/Huaray/                  → Camera SDK DLLs (copied local)
├─ dll/                                  → SolVision native support DLLs
├─ images/                               → UI icons / samples
├─ bin_temp/                             → scratch output dir used during builds (git-tracked for convenience)
└─ ...                                   → other assets noted in the project file
```

> **Note**: `bin/`, `obj/`, `Captured/`, and other transient build folders are ignored via `.gitignore`.

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| **OS** | Windows 10/11 x64 |
| **IDE** | Visual Studio 2022 (17.6+) with `.NET desktop development` workload |
| **.NET** | .NET Framework 4.8 SDK (installed with VS desktop workload) |
| **SDKs** | SolVision 6.1.4 runtime, Huaray MV SDK (USB & GigE drivers), ComXim turntable USB driver |
| **Git** | Git + Git LFS (required if you clone the binary repo later) |
| **Python** | Bundled `python38.dll` for SolVision Python interop (already in `dll/`) |
| **GPU (optional)** | CUDA-capable GPU accelerates SolVision inference |

Install vendor SDKs **before** running within Visual Studio to avoid missing DLL runtime errors.

---

## Building & Running

1. Clone the repo and open `DemoApp.sln` in Visual Studio.
2. Select build configuration:
   - `Release | x64` for production (matches the stable binary build).  
   - `Debug | x64` for development.
3. Build (`Build → Build Solution`).
4. Run (`Debug → Start Debugging` or `F5`). Ensure your hardware is connected.

### Optional command-line build

```powershell
msbuild DemoApp.sln /p:Configuration=Release /p:Platform=x64
```

> Output goes to `DemoApp/bin/x64/<Config>/`. After testing, copy `Release` into `bin\x64\Stable` before packaging.

---

## Hardware Setup

1. **SolVision Licensing**: ensure the machine has a valid SolVision license.  
2. **Turntable**: connect via USB (COM port). Install ComXim driver; note COM number.  
3. **Cameras**:
   - USB top camera (Huaray).  
   - GigE front camera (Huaray). Ensure NIC and camera share same subnet (e.g., 169.254.x.x).  
   - Use Huaray “MV Viewer” to confirm the cameras stream before running the app.

4. **GPU**: optional but recommended. Ensure NVIDIA drivers are current if using CUDA.

---

## Workflow Walkthrough

### Initialize Tab

1. **Load Projects**  
   - Attachment `.tsp` → loads into `attachmentContext`.  
   - Defect `.tsp` → loads into `defectContext`.  
   Status colors (neutral / success / warning) reflect model loading.

2. **Cameras**  
   - Refresh device list (top + front).  
   - Connect each; capture a preview to validate.  
   - Preview images route to the workflow panel for reference.

3. **Turntable**  
   - Select COM port, connect, then home.  
   - Homing logs the offset angle; the turntable must be homed before detection is enabled.

4. **Summary Banner**  
   - Displays readiness across models, cameras, and turntable.  
   - Turns green when the system is ready to inspect.

### Workflow Tab

1. **Detect**  
   - Captures the top image.  
   - Runs the attachment model; draws overlay (center marker + sequence numbers).  
   - Stores `AttachmentPointInfo` list with angles, center coordinates, and sequence indexes.

2. **Automatic Capture Sequence**  
   - For each attachment point:  
     - Rotate the turntable (CW or CCW depending on orientation).  
     - Capture front image; save to `Captured/`.  
     - When returning home, process images with the **defect** model.

3. **Front Inspections UI**  
   - Gallery of inspection cards (PASS/FAIL badges).  
   - Dual preview: top overlay (left) and selected front overlay (right).  
   - Ledger table: index, class, confidence, area, bounding box.

4. **Status Banner**  
   - Shows PASS/FAIL after logic evaluation (see next section).

---

## Logic Builder

- Rules now evaluate **defect detections** (front inspections) rather than attachment results.
- Fields:
  - `Class Name`, `Confidence`, `Area`, `Count` (auto-suggested values update after each inspection).
- Rule groups can be `ALL` (AND) or `ANY` (OR).
- If **any** child rule/group evaluates to `TRUE`, the inspection fails (red banner, log entry).
- `?` button shows users a help dialog summarizing logic semantics.

Implementation highlights:

```csharp
bool fail = EvaluateNode(logicRoot, CollectCurrentDefectDetections(), out LogicNodeBase triggered);
```

- `CollectCurrentDefectDetections()` aggregates detection results across all front inspection results.
- `FrontInspectionResult` tracks summary text, overlays, and detection metadata.

---

## Turntable & Cameras

- `TurntableController` in `DemoApp.cs` handles serial protocol (`CT+START`, `CT+CHANGESPEED`, `CT+GETOFFSETANGLE`).
- Speed adjustments (for quicker cycle time):
  ```csharp
  turntableController.SendCommand("CT+CHANGESPEED(1);");
  turntableController.SendCommand("CT+GETTBSPEED();"); // Logs current/max
  ```
- `CameraContext` wraps Huaray device handles, capturing frames via `MVSDK_Net.dll`.
- `CaptureCameraFrame()` returns `Bitmap` for preview and saved PNGs.
- Angle math uses `angle_calculation.py` logic ported to C# (`NormalizeSignedAngle`, center-based calculations).

---

## Coding Guidelines

- **Target Framework**: .NET Framework 4.8; avoid API calls requiring .NET 5+.
- **Style**: match existing patterns (PascalCase methods, camelCase locals).  
- **Threading**: heavy operations (detection, turntable moves) run on background threads; UI updates must call `BeginInvoke`.
- **Logging**: use `outToLog()` with `LogStatus` enum to keep operator feedback consistent.
- **Resource Management**: dispose `Bitmap`, `Mat`, and `FrontInspectionResult` objects to prevent memory leaks.
- **Error Handling**: user-facing warnings should explain corrective actions (e.g., “Connect front camera”).
- **Testing**: manual regression across detection scenarios (pass/fail, missing hardware). Automated tests are not yet present; consider adding for future.

---

## AI Assistant Usage

AI coding tools (GitHub Copilot, ChatGPT, etc.) are welcome, but follow these guardrails:

1. **Stay in this repo** – always modify source here, never the binary release repo.
2. **Explain changes** – when using AI to generate code, review carefully, add comments if logic is complex, and document assumptions.
3. **Protect hardware safety** – review any change touching turntable motion, camera drivers, or asynchronous operations.
4. **Formatting & style** – keep coding style consistent with existing files; run `Format Document` (Ctrl+K, Ctrl+D) in VS.
5. **Commit hygiene** – use descriptive messages (e.g., `Improve defect gallery spacing`) and push via pull requests when collaborating.

If AI suggests large refactors, discuss with the team before landing to avoid breaking workflows.

---

## Release Process

1. **Develop** & test changes on a feature branch.
2. **Merge** into `main` (or designated release branch).
3. **Build** `Release|x64` → verify full workflow with real hardware.
4. **Copy** `DemoApp/bin/x64/Release` to `DemoApp/bin/x64/Stable`.
5. **Package** the stable folder into the binary repo:
   ```powershell
   Copy-Item .\DemoApp\bin\x64\Stable\* ..\Envista_turntable_stable\ -Recurse -Force
   ```
6. **Update** the binary repo README if behavior changed, commit, and push.
7. Optionally, create a GitHub Release/tag in both repos for tracking.

---

## Troubleshooting

| Issue | Resolution |
|-------|------------|
| `System.IO.FileNotFoundException: SolVision.dll` | Install/run SolVision installer; ensure `dll/` contents copy next to the EXE or fix probing path. |
| MVSDK DLL missing | Reinstall the Huaray MV SDK; confirm `MVSDKmd.dll` ships in `PATH` or `Dependencies/Huaray/`. |
| Turntable timeouts (`CR+ERR` or no TB_END) | Check COM port, ensure no other app controls the turntable, and confirm `CT+GETOFFSETANGLE` returns data. |
| App crashes on detection | Validate TSP project compatibility (object/class names). Ensure GPU drivers/current license. |
| Git push rejected (large file) | Use Git LFS if adding large assets, or exclude heavy binaries from source repo. |

For escalations, contact the Desktop Vision/Envista inspection team or the hardware vendors.

---

Happy coding! 💡 Let the team know via Slack when you push significant changes so the stable build can be regenerated. Remember: source changes happen here; production binaries live in `Envista_turntable_stable`.
