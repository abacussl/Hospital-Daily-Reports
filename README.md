🏥 Hospital Day Report Dashboard (Google Apps Script Project)

This repository contains a Google Apps Script + HTML web app for collecting, managing, and visualizing daily hospital activity data across multiple departments.

📌 Overview

The system allows hospital staff to enter daily statistics from different units — OPD, Blood Bank, Hematology, Maternal Census, Dialysis, Theaters, and more — into a centralized dashboard. The data is stored in Google Sheets via Apps Script and can be reviewed or analyzed later.

Key features:

🗓 Date-based reporting (choose a date to view or submit data)

🏥 Multi-department forms with validation

📊 Dynamic UI: buttons turn green once data is submitted for a section

☁️ Serverless architecture using Google Apps Script and HTML Service

📱 Responsive design powered by Bootstrap 5

📂 Project Structure
.
├── code.gs        # Google Apps Script backend functions
└── index.html     # HTML frontend with Bootstrap forms and client-side JS

Backend (code.gs)

Handles CRUD operations on Google Sheets

Provides getFormData() to fetch saved data for a given date

Provides submitXXX() functions for each section (OPD, Blood Bank, Hematology, etc.)

Returns success responses to update UI status indicators

Frontend (index.html)

Uses Bootstrap 5 for layout and styling

Dynamically loads forms for each section

Pre-fills fields with saved data (if available)

Submits form data using google.script.run to call Apps Script functions

Shows visual feedback (alerts + button highlighting)

🖥️ Features by Section

OPD Report – Visits, nebulizations, injections, dressings, minor surgeries

Endoscopy Unit – Total endoscopies performed

Blood Bank – Blood/component issues, cross-matching, collections, stock

Hematology – Admissions, phlebotomy, platelet counts, referrals

Blood Drawing Centers – Room-wise collection counts

Infectious Diseases – Dengue, COVID-19, Influenza B, etc.

Pain Management – Post-op pain assessments, clinic follow-ups

Chemotherapy Unit – Chemotherapy, zoledronic acid, filgrastim treatments

Breastfeeding Day Centre – Referrals and educational sessions

Maternal Census – Deliveries, stillbirths, neonatal/maternal deaths

Dialysis Unit – Dialysis counts, CAPD inserts, HD catheter data

A&E Midnight Summary – Admissions, discharges, transfers, deaths

Operating Theaters (AB, CD, EF, GH, M, Eye) – Major/minor surgeries, anesthesia breakdown

🚀 Deployment

Open Google Apps Script
.

Create a new project and copy the contents of code.gs into the script editor.

Add a new HTML file and paste the contents of index.html.

Link to a Google Sheet (ensure correct sheet names and ranges).

Deploy as a Web App:

Go to Deploy → New Deployment

Set "Web app" type

Choose "Anyone with link" or restrict to internal users

Copy and share the web app URL

🔧 Technologies Used

Google Apps Script (Server-side)

Google Sheets (Database)

HTML, JavaScript

Bootstrap 5 (Responsive UI)

📊 Example Usage

Staff selects the date

Fills out the relevant department forms

Submits data → Button turns green (visual confirmation)

Admin can later review data in linked Google Sheet

🤝 Contributing

Pull requests are welcome!
If you want to add new sections or improve data validation, feel free to fork and submit a PR.

📜 License

This project is licensed under the MIT License – feel free to use and modify for your hospital’s reporting needs.
