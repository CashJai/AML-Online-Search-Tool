# AML Online Search Tool

A high-precision AML (Anti-Money Laundering) and Adverse News screening engine designed to automate subject screening while significantly reducing false positives through advanced name-matching heuristics and multi-category FCC (Financial Crime Compliance) detection.

## 🚀 Overview
This tool was built to streamline the workflow for AML analysts. It performs deep-web searches (via Bing) to identify potential matches for subjects across general online presence and specific adverse news categories.

### Why this is different
Most standard search tools suffer from "Keyword Noise." This tool solves that by using **deterministic rule-based logic** instead of unpredictable LLMs, ensuring that every "True Hit" is auditable and based on strict linguistic boundaries.

## 🧠 Key Logic & Features

### 1. Advanced Name-Matching Engine
The tool doesn't just look for substrings. It uses a custom-built matching engine that handles:
*   **Cartesian Alias Generation:** Automatically expands subjects into valid linguistic permutations (e.g., "Chan" ↔ "Chen") using a custom fuzzy dictionary.
*   **Strict Boundary Validation:** Eliminates false positives like "Chan Tai Man" when searching for "Chan Tai" by analyzing capitalized neighbors.
*   **Leading Name Logic:** Correcty identifies variations like "David Chan Tai" as valid hits while rejecting trailing name overlaps.

### 2. Robust FCC Concern Detection
The Adverse News engine uses regex-based stemming to identify 20+ categories of financial crime, including:
*   **Corruption & Bribery**
*   **Money Laundering & Fraud**
*   **Sanctions & Terrorist Financing**
*   **Litigation & Criminal Proceedings**

### 3. Multilingual Support (English & Chinese)
Fully supports screening in both English and Chinese, with automated category detection for Simplified and Traditional Chinese characters.

### 4. Professional Reporting
Generates comprehensive, audit-ready reports:
*   **Excel:** Consolidated multi-tab workbooks with summary headers, risk levels, and direct source links.
*   **PDF:** Full-page evidence captures for compliance record-keeping.

## 🛠 Tech Stack
*   **Language:** Node.js
*   **Automation:** Playwright (Chromium)
*   **Data Processing:** XLSX-JS, PDF-Lib
*   **Linguistic Processing:** Chinese-Conv, Custom Regex Stemming

## 📂 Project Structure
*   `Online Search Tool - Mac/`: optimized for macOS users.
*   `Online Search Tool - Windows/`: optimized for Windows users.

---
**Author:** Marcus Leung  
**Contact:** [cl4142@columbia.edu](mailto:cl4142@columbia.edu)
