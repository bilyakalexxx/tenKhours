# ⏳ tenKhours

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python Version" />
  <img src="https://img.shields.io/badge/UI-CustomTkinter-2b5c8f?style=for-the-badge" alt="CustomTkinter" />
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Platform Windows" />
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="License MIT" />
</p>

<p align="center">
  <b>A minimalist, automated desktop tracker engineered to help creatives, software engineers, and digital artists master their crafts—one focused hour at a time.</b>
</p>

---

## 📌 Overview

The **10,000-Hour Rule** states that deep mastery in any complex discipline requires roughly 10,000 hours of deliberate practice. **tenKhours** is a Windows desktop application that runs in the background, automatically tracking the active focus time spent inside specific software environments.

Instead of generic time trackers that measure raw screen time, **tenKhours** groups your tools into **Group Containers** (e.g., *3D Modeling*, *Game Development*, *Full-Stack Coding*) and measures your individual and cumulative progress toward total mastery.

---

## ✨ Key Features

- 🎯 **Automated Window Tracking**: Detects active focus windows on Windows OS using low-level API hooks (`pywin32` + `psutil`).
- 📁 **Group Containers**: Bundle multi-application workflows (e.g., Blender + Maya + ZBrush under a single "3D Art" container).
- 📊 **Per-Program Progress Tracking**: Real-time progress bars and hour counts for every assigned `.exe` relative to your 10,000-hour goal.
- ⏱️ **Zero-Friction Live Stopwatch**: Displays precise per-session tracking and automatically updates overall group statistics in real time.
- 💾 **SQLite Storage & Persistence**: Local, crash-resilient SQLite logging ensures no lost focus time.
- 📈 **Excel Report Exports**: One-click formatting export (`.xlsx`) to review historical focus sessions with pre-formatted durations.
- 🌙 **Modern Dark Theme**: Native dark interface built on `CustomTkinter`.

---

## 🛠️ Tech Stack & Dependencies

- **GUI Framework**: [`CustomTkinter`](https://github.com/TomSchimansky/CustomTkinter)
- **OS Integration**: `pywin32`, `psutil`
- **Data & Export**: `sqlite3`, `pandas`, `openpyxl`
- **Language**: Python 3.8+

---

## 📖 App Guide

![Guide Step 1](guide/guide_01.png)
![Guide Step 2](guide/guide_02.png)
![Guide Step 3](guide/guide_03.png)
![Guide Step 4](guide/guide_04.png)
![Guide Step 5](guide/guide_05.png)