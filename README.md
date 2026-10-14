# High-Performance 24-Hour Timeline Control for WinCC V8

A robust, multi-layered industrial automation solution designed to monitor and visualize 24-hour operational status timelines for industrial machines (e.g., Filter Presses) in **Siemens WinCC V8**, bypassing traditional SCADA limitations using **Custom Web Controls (HTML5/CSS3/JS)** and **OLEDB Tag Logging**.

---

## 🏗️ System Architecture & Data Flow

The project follows a 3-tier hybrid ETL architecture to bridge WinCC historical archives with modern web-based UI components:

1. **Extraction (VBScript & OLEDB):** Upon opening the HMI screen (`OnOpen`), an embedded VBScript connects to the WinCC runtime database (`@DatasourceNameRT`) via `WinCCOLEDBProvider.1` and queries 24–48 hours of historical tag data (`TAG:R`).
2. **Transformation & Tokenization:** The script compresses state changes into compact token strings (Mode ID + Timestamp) and distributes them across internal WinCC text tags (`_ModeData1` to `_ModeData4`) to overcome character length limits.
3. **Frontend Rendering (Custom Web Control):** The HTML/JavaScript engine parses the text tags, calculates time proportions, applies `ResizeObserver` for dynamic responsive scaling, and renders a sleek, industrial dark-themed timeline with active duration timers.

---

## ⚙️ How to Configure in WinCC V8

1. **Tag Management:** Create internal text tags (`_ModeData1` to `_ModeData4`) and process tags for your machines (`FilterPress1_WTC_Mode`, etc.).
2. **Tag Logging:** Set up process value archives with an acquisition type of *After every change*.
3. **VBScript:** Attach the script from `vbs-scripts/OnOpen_Timeline_ETL.vbs` to the `OnOpen` event of your Graphics Designer screen.
4. **Custom Web Control:** Place the web control component on the screen and link its properties (`status`, `modeData1` through `modeData4`) to the corresponding WinCC tags.

---

## 📄 Documentation

For detailed technical breakdowns, code interpretations, and troubleshooting guides, please check the [`docs/scada_timeline_manual_en.pdf`](docs/scada_timeline_manual_en.pdf) file in this repository.

---

## 📂 Repository Structure

```text
WinCC-Custom-Timeline-Widget/
├── README.md                              # Project documentation
├── docs/
│   └── scada_timeline_manual_en.pdf       # Comprehensive technical manual (PDF)
├── vbs-scripts/
│   └── OnOpen_Timeline_ETL.vbs            # Optimized VBScript for WinCC OnOpen event
└── web-control/
    ├── index.html                         # Frontend HTML & JavaScript parsing logic
    ├── style.css                          # Industrial dark-theme styling
    └── webcc.min.js                       # Siemens WebCC library
