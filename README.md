# Automated Dispatch & Analytics Engine (Excel/VBA)

## 📌 Overview
This project is a fully automated, scalable resource management and analytics tool built entirely within Excel using advanced VBA, dynamic arrays, and REST APIs. 

Originally developed to manage logistics, performance tracking, and communications for a 50+ member gaming organization, the system functions as a lightweight ETL (Extract, Transform, Load) pipeline. Featuring a responsive UI with locked navigation headers, the tool automatically parses user performance data, updates a relational database, sanitizes backend logic for public reporting, and dispatches automated JSON payloads to Discord via webhooks.

## 🚀 The Problem
Managing daily operations for a large-scale team required hours of manual data entry, performance tracking, and communication. Specifically:
* **Logistics:** Manually assigning tasks and tracking real-time availability.
* **Analytics:** Calculating historical user accuracy and determining status upgrades based on strict performance criteria.
* **Security & Reporting:** Sharing progress data with the team without exposing proprietary backend formulas, sensitive internal metrics, or administrative notes.
* **Communications:** Pinging dozens of users with their specific assignments without triggering platform rate limits or spam filters.

## 💡 The Solution
A comprehensive Excel-based management tool that reduces hours of administrative overhead to a single button click. The system utilizes VBA macros to execute automated data cloning, formula flattening, bulk assignments, and delayed API broadcasts.

## ✨ Core Features

### 1. Automated ETL & Data Sanitization
* **One-Click Publishing:** A VBA macro instantly duplicates active data sheets and generates a separate, public-facing Excel workbook.
* **Formula Flattening:** Automatically strips out structured table arrays and converts complex background formulas into static text to protect backend architecture and ensure file stability.
* **Data Redaction:** Programmatically hides specified columns containing sensitive administrative data or future projections before saving.
* **Automated Cleanup:** Iterates through the generated workbook to delete obsolete tabs, ensuring the public file remains accurate and clean.

### 2. Custom API Integration (Discord Webhooks)
* **Automated Dispatch:** Utilizes `WinHttp.WinHttpRequest.5.1` to format dynamic Excel data into JSON payloads and POST them directly to Discord webhooks.
* **Logic-Based Sorting:** Scans the active roster, groups assignments by priority, and physically sorts lower-priority data to the bottom of the transmission list.
* **Payload Batching & Rate-Limit Handling:** Groups targeted `@mentions` into perfectly sized blocks (e.g., 5 per message) and injects a 1-second `Application.Wait` with a `DoEvents` failsafe to bypass server anti-spam filters without crashing the Excel UI.
* **Scheduled Relays:** Utilizes `Application.OnTime` to send a follow-up ping to specific server roles exactly 3 minutes after the initial broadcast.

### 3. Dynamic KPI Tracking & Relational Database
* **Automated Data Fetching:** Uses dynamic array formulas (`UNIQUE`, `FILTER`) to automatically generate clean member lists from raw log data.
* **Condition-Based Status Updates:** Automatically tracks user performance across multiple events. If a user hits strict logical parameters (e.g., executing high-value targets with perfect scores across 3 events), they are automatically indexed into the qualified relational database.
* **Live Analytics:** Cross-references the active roster with the database using `XLOOKUP` to display current capabilities and calculates historical accuracy metrics using nested `COUNTIFS`.

### 4. Scalable Bulk Operations
* **Responsive Architecture:** The template relies on dynamic ranges, requiring zero structural changes whether managing 15 users or 150 users.
* **Bulk Assignment Automation:** A built-in macro reads the active timeframe, utilizes Unicode injection to bypass legacy VBA character limitations, and automatically populates the daily roster with default assignments based on interactive user prompts.

## 📸 System Previews

(to be added)

## 🛠️ Technology Stack
* **Advanced Excel:** Dynamic Arrays, Complex Logic Functions, Conditional Formatting, Data Validation.
* **VBA (Visual Basic for Applications):** Macro Automation, Windows File System Object, HTTP Requests, Memory Management.
* **Integrations:** REST APIs, JSON, Discord Webhooks.
