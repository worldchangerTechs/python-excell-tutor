# 📊 ExcelTutor AI

<div align="center">

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![NVIDIA NIM](https://img.shields.io/badge/AI%20Engine-NVIDIA%20NIM-76B900.svg?logo=nvidia&logoColor=white)](https://build.nvidia.com/)
[![Excel Automation](https://img.shields.io/badge/Automation-xlwings%20%7C%20COM-217346.svg?logo=microsoftexcel&logoColor=white)](https://www.xlwings.org/)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-0078D6.svg?logo=windows&logoColor=white)](https://www.microsoft.com/)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/)

**The autonomous, voice-driven AI tutor and automation engine that controls your desktop Excel in real time.**

*Not another copy-paste chatbot. It watches your active workbook, executes real formulas, builds dashboards, draws charts, and autoscrolls to teach you.*

[Key Features](#-key-features) • [Quickstart](#-quickstart) • [How It Works](#-how-it-works) • [Demo Prompts](#-demo-prompts) • [Architecture](#-architecture) • [Contributing](#-contributing)

</div>

---

## 💡 Why ExcelTutor AI?

Most "Excel AI" tools are just text chatbots: they output formulas in a markdown box and leave you to copy, paste, debug, and format everything yourself.

**ExcelTutor AI is an agentic co-pilot.** It connects directly to your live, running Microsoft Excel desktop application via Windows COM automation. When you ask a question or request an analysis:
1. It inspects your real spreadsheet data and active cell selection.
2. It **directly writes formulas, builds tables, applies heatmaps, and draws charts** in your open workbook.
3. It **autoscrolls and highlights cells in motion**, teaching you step-by-step why formulas work.
4. It **speaks back to you via text-to-speech** while keeping explanations concise and focused on the screen.

```
                  ┌────────────────────────────────────────┐
                  │          🎙️ Voice / ⌨️ Text           │
                  │   "Teach me data mining on this sheet" │
                  └───────────────────┬────────────────────┘
                                      │
                                      ▼
                  ┌────────────────────────────────────────┐
                  │       🧠 AI Brain (NVIDIA NIM)        │
                  │   High-speed inference + Auto-fallback │
                  └───────────────────┬────────────────────┘
                                      │
                         Structured Action Pipeline
                                      │
                                      ▼
                  ┌────────────────────────────────────────┐
                  │     ⚡ Real-Time Windows COM Engine     │
                  │  xlwings • PyAutoGUI • Threaded COM    │
                  └───────────────────┬────────────────────┘
                                      │
          ┌───────────────────────────┴───────────────────────────┐
          ▼                                                       ▼
┌───────────────────────────────┐               ┌───────────────────────────────┐
│       🟢 Active Excel App     │               │     💻 Floating UI Tutor      │
│ • Live Formula Injections     │               │ • Always-on-top Glass UI     │
│ • Instant 3-Color Heatmaps    │               │ • Real-time Speech-to-Text    │
│ • Native Excel ListObjects    │               │ • Natural Voice Synthesis     │
│ • Smooth Screen Autoscrolling │               │ • Kernel-Level Anti-Sleep     │
└───────────────────────────────┘               └───────────────────────────────┘
```

---

## ✨ Key Features

### ⚡ 1. In-Sheet Real-Time Execution
- **Zero Copy-Pasting:** Changes happen inside your `.xlsx` file as you watch.
- **Dynamic Formula Builder:** Injects live Excel formulas (`=SUM()`, `=AVERAGE()`, `=VLOOKUP()`, `=STDEV.S()`, `=COUNTIF()`).
- **Preserves Your Data:** Intelligently reads existing sheets without overwriting your original raw records.

### ⛏️ 2. Autonomous Data Mining & Profiling
- **One-Click Mining:** Click **`⛏️ Mine Data`** or ask by voice to profile any dataset.
- **Automated Summary Dashboards:** Builds complete descriptive statistics tables (Average, Median, Min, Max, Standard Deviation) alongside your data.
- **Categorical Breakdown:** Automatically generates frequency tables and distribution metrics.
- **Pattern Heatmaps:** Injects native 3-color conditional formatting scales to reveal data density, clusters, and trends.

### 📊 3. Native Charts & Tables
- **Formatted Tables:** Instantly converts raw data ranges into structured Excel `ListObjects` with custom table styles.
- **Chart Generation:** Creates Column, Bar, Line, Pie, Area, and Scatter plots positioned cleanly next to your data.
- **Auto-Fit Formatting:** Automatically calculates and sets optimal column widths and header styles.

### 🎙️ 4. Full Voice Interaction
- **Hands-Free Tutoring:** Tap **`🎤 Speak`** and talk naturally to your spreadsheet.
- **Auditory Feedback:** Integrated offline Text-to-Speech (TTS) engine explains what it's doing so you can keep your eyes on the sheet.

### 🛡️ 5. Kernel-Level Anti-Sleep Engine
- **Presentation Ready:** Powered by Windows `SetThreadExecutionState` (`ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED`).
- Prevents your computer from sleeping, hibernating, or dimming the display during long presentations, tutoring sessions, or live demos.
- Automatically restores standard power policies on exit.

### 🧠 6. Enterprise-Grade AI with Auto-Fallback
- Powered by high-throughput **NVIDIA NIM** microservices (`mistral-large`, `nemotron-3.5-lightning`, `minimax-m3.1`).
- Built-in multi-model resilience: automatically catches deprecated endpoints or rate limits and fails over with zero user disruption.

---

## 🚀 Quickstart

### Prerequisites
- **Operating System:** Windows 10 or Windows 11 (with Microsoft Excel installed)
- **Python:** Python 3.10, 3.11, or 3.12

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/excel-tutor-ai.git
cd excel-tutor-ai
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```
*(Or install manually: `pip install openai xlwings pyautogui SpeechRecognition pyttsx3 python-dotenv`)*

> **Optional for microphone input:** Install PyAudio:
> ```bash
> pip install pyaudio
> ```

### 3. Configure Your Environment
Create or edit your `.env` file in the project root:
```env
NVIDIA_API_KEY=your_nvidia_nim_api_key_here
NVIDIA_MODEL=mistralai/mistral-large-4-1t-a52b
```

### 4. Launch ExcelTutor AI
```bash
python app.py
```
*Excel will automatically open (or attach to your existing workbook), and the tutor panel will float on top ready for instructions!*

---

## 🎮 Demo Prompts

Try saying or typing any of these commands:

| Goal | Prompt | What It Does in Excel |
|---|---|---|
| **Data Mining** | `"Mine this data and show me key patterns"` | Builds a statistical profiling table, applies a heatmap, and summarizes findings. |
| **Outlier Detection** | `"Find outliers in credit limit and highlight them"` | Scans distribution, flags anomalies (>2 std dev) in orange/red fill. |
| **Formulas & Math** | `"Calculate monthly profit margin in a new column"` | Injects formula into the first row and fills it down the column. |
| **Interactive Tour** | `"Walk me through this sheet and explain the headers"` | Highlights columns one by one and autoscrolls through records. |
| **Visualization** | `"Create a column chart comparing sales by region"` | Identifies categories, creates aggregation, and inserts a styled chart. |
| **Exporting** | `"Export this cleaned sheet to CSV on my Desktop"` | Generates a clean UTF-8 encoded `.csv` file directly. |

---

## 🏗️ Architecture

```
excel-tutor-ai/
│
├── app.py             # Main application: Tkinter GUI, Windows COM loop, AI prompt engine
├── .env               # API credentials & model configurations
├── requirements.txt   # Python package dependencies
└── README.md          # Project documentation & guides
```

### Technical Highlights:
- **Apartment-Threaded COM:** Background worker threads safely initialize Windows COM via `pythoncom.CoInitialize()` and `pythoncom.CoUninitialize()`, ensuring Excel never freezes the GUI.
- **Edit-Mode Resilience:** Detects active workbooks and skips conflicting locks while handling large data frames through 2D array blocks rather than slow cell-by-cell loops.
- **Fail-Safe Navigation:** Built-in `pyautogui.FAILSAFE = True` lets you abort automated cursor movements at any time by flicking your mouse to any screen corner.

---

## 🛠️ Supported Action Primitives

The AI tutor engine translates natural language into structured automation actions:

```json
{
  "say": "I've added the data mining dashboard to column Y and applied a 3-color heatmap.",
  "actions": [
    { "type": "write_data", "start_cell": "Y4", "data": [["Metric", "Avg", "Max"], ["Age", "=AVERAGE(C2:C1000)", "=MAX(C2:C1000)"]] },
    { "type": "highlight", "range": "Y4:AA4", "color": "blue" },
    { "type": "colorscale", "range": "Z5:AA15" },
    { "type": "autofit" },
    { "type": "goto", "cell": "Y1" }
  ]
}
```

---

## 🤝 Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

---

<div align="center">

**⭐ Love this project? Leave a star on GitHub to support development! ⭐**

*Built with ❤️ for data analysts, finance professionals, and Excel learners.*

</div>
