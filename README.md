# Automated Dispatch & Analytics Engine (Excel/VBA)

## 📌 Overview
This project is a fully automated, highly scalable resource management and analytics engine built entirely within Excel using advanced VBA, dynamic arrays, and REST APIs. 

Originally engineered to manage logistics, performance tracking, and communications for a 50+ member gaming organization (Clan War Leagues), the system functions as a lightweight ETL (Extract, Transform, Load) pipeline. It automatically parses user performance data, updates a relational database, sanitizes backend logic for public reporting, and dispatches automated JSON payloads to Discord via webhooks.

## 🚀 The Problem
Managing daily operations for a large-scale team required hours of manual data entry, performance tracking, and communication. Specifically:
* **Logistics:** Manually assigning targets and tracking real-time availability.
* **Analytics:** Calculating historical user accuracy and determining promotions based on strict performance criteria.
* **Security & Reporting:** Sharing tactical data with the team without exposing proprietary backend formulas, sensitive internal metrics, or administrative notes.
* **Communications:** Pinging dozens of users with their specific assignments without triggering platform rate limits or spam filters.

## 💡 The Solution
An "overengineered" Excel-based command center that reduces hours of administrative overhead to a single button click. The system utilizes VBA macros to execute data cloning, logic disarming, conditional mass-assignments, and delayed API broadcasts.

## ✨ Core Features

### 1. Automated ETL & Data Sanitization (The "Decoy" Dashboard)
* **One-Click Publishing:** A VBA macro instantly duplicates active data sheets and generates a separate, public-facing Excel workbook.
* **Logic Disarming:** Automatically strips out structured table arrays and forcefully converts complex background formulas into pure, static text to protect backend architecture.
* **Tactical Redaction:** Programmatically hides specified columns containing sensitive administrative data or future projections before saving.
* **Self-Cleaning:** Iterates through the workbook to delete obsolete tabs, ensuring the public file remains pristine.

### 2. Custom API Integration (Discord Webhooks)
* **Automated Dispatch:** Utilizes `WinHttp.WinHttpRequest.5.1` to format dynamic Excel data into JSON payloads and POST them directly to Discord webhooks.
* **Logic-Based Sorting:** Scans the active roster, groups assignments by priority, and physically sorts lower-priority data to the bottom of the transmission list.
* **Chunking & Rate-Limit Evasion:** Groups targeted `&mentions` into perfectly sized blocks (e.g., 5 per message) and injects a 1-second `Application.Wait` with a `DoEvents` failsafe to bypass server anti-spam filters without crashing the UI.
* **Time-Delayed Relays:** Deploys a secondary macro via `Application.OnTime` to send a follow-up ping to specific server roles exactly 3 minutes after the initial broadcast.

### 3. Dynamic KPI Tracking & Relational Database
* **Automated Roster Fetching:** Uses dynamic array formulas (`UNIQUE`, `FILTER`) to automatically generate clean member lists from raw log data.
* **Condition-Based Promotions:** Automatically tracks user performance across multiple events. If a user hits strict logical parameters (e.g., executing high-value targets with perfect scores 3 times), they are automatically indexed into the "Qualified" relational database.
* **Live Analytics:** Cross-references the active roster with the database using `XLOOKUP` to display current power levels and calculates historical accuracy win-rates using nested `COUNTIFS`.

### 4. Scalable UI & Command Deck
* **Responsive Architecture:** The template relies on dynamic ranges, meaning it requires zero structural changes whether managing 15 users or 150 users.
* **Frozen Command Deck:** Key macros, day-trackers, and execution buttons are locked to the top of the screen for rapid access.
* **Mass Assignment Drone:** A macro prompt that reads the active timeframe, utilizes Unicode injection to bypass legacy VBA character limitations, and automatically populates the daily roster with default strategic assignments. 

## 📸 System Previews

(to be added)

## 🛠️ Technology Stack
* **Advanced Excel:** Dynamic Arrays, Complex Logic Functions, Conditional Formatting, Data Validation.
* **VBA (Visual Basic for Applications):** Macro Automation, Windows File System Object, HTTP Requests, Memory Management.
* **Integrations:** REST APIs, JSON, Discord Webhooks.
