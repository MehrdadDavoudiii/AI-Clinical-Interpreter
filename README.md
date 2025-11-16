# AI Clinical Interpreter

AI Clinical Interpreter is a free, open-source desktop tool designed to help clinicians, researchers, and students generate structured clinical reports from anonymized gait and motion analysis data and other orthopedic / movement-related assessments.

The tool sends the content of selected anonymized files (e.g., PDF, DOCX, TXT, images) to a large language model (e.g., GPT-5.1) and produces a Word document (.docx) that combines a local patient header (kept only on your computer) with an AI-generated clinical interpretation of the anonymized data.

---

## 1) Overview

This tool provides an intuitive interface to:

- Enter a local patient header (kept only on your computer).
- Select a save location for the final report.
- Load anonymized clinical files (PDF, DOCX, TXT, images).
- Generate a structured clinical report using an AI model via the OpenAI API.
- Save a formatted, editable Word report (.docx).

Only the anonymized content from the selected files is sent to the AI model. All patient header information remains local and is never transmitted.

---

## 2) Download & Usage (Portable App)

This application can be provided as a portable .zip file. No installation is required for the portable version.

**Download the .zip file:**

- Obtain the .zip archive from the distribution link provided by the developer (e.g., university file share, GitHub release, or institutional download page).

**Extract the .zip file:**

- Right-click the downloaded .zip file and select “Extract All…”.
- Choose a folder to extract the application. It is recommended to use a simple path on a main drive (e.g., `C:\AI_Clinical_Interpreter\`) rather than your “Downloads” folder.

**Run the application:**

- Open the folder you just extracted.
- Double-click the `AI_Clinical_Interpreter.exe` file to run the application.

(If you are running from source instead of the .exe, activate your Python environment and run the main `.py` file.)

---

## 3) Key Features

- **Structured Clinical Reports**  
  Generates a concise, structured Word report (.docx) designed for busy clinicians (surgeons, physiotherapists, rehabilitation specialists).

- **Local Patient Header**  
  Patient name, ID, birthdate, and other header fields are entered locally and never sent to the AI. They are merged into the final report only on your machine.

- **Multi-Format Input**  
  Supports multiple file types for anonymized input:
  - PDF (e.g., full gait reports, lab exports)
  - DOCX (clinical notes, summaries)
  - TXT (raw values, simple exports)
  - Images (e.g., screenshots of plots or tables)

- **Integrated Clinical Prompt**  
  Uses a dedicated, clinically oriented prompt to combine multiple data sources into a single, decision-oriented narrative with key abnormalities and recommended next steps.

- **Simple Graphical Interface**  
  Built with Tkinter (and optional modern theming) for a clean, desktop-friendly workflow.

---

## 4) System Requirements

- Windows 10 or later (for the compiled .exe version).
- Internet connection for calling the OpenAI API.

For running from source:

- Python 3.10 or later (recommended).
- Required Python packages (examples):
  - `openai`
  - `python-docx`
  - `pymupdf` (PyMuPDF)
  - `pillow`
  - `sv-ttk` (optional theming)

---

## 5) Quick Start

1. Launch the application by double-clicking `AI_Clinical_Interpreter.exe`.
2. Fill in the **Patient Header** fields (e.g., first name, last name, patient ID, birthdate).
   - This header remains local and is not sent to the AI.
3. Select the **Save Location** for the final report.
   - The application will propose a default filename (e.g., `Report_<LastName>_<PatientID>.docx`), which you can adjust.
4. Click **Add File** to load one or more anonymized clinical documents (PDF, DOCX, TXT, images).
   - For each file, provide a short description (e.g., “Spatiotemporal parameters”, “Surface EMG”, “Lumbar MRI summary”).
5. When you are ready, click the **Process** / **Finish & Process** button.
6. Confirm the anonymization and privacy notice.
7. Enter your **OpenAI API key** when prompted.
   - The key is used to contact the AI model and is not stored permanently by default.
8. Wait while the status bar shows progress (file reading, AI interpretation, report generation).
9. When finished, you will be notified and asked if you want to open the output folder containing the generated `.docx` report.

---

## 6) Help Guide

A detailed user guide can be integrated into the application.

- Launch the app and click **Help** > **Usage Help** (or a similar menu item) in the menu bar to read the built-in help text explaining each step and option.

---

## 7) Data Storage

- Generated reports are saved as `.docx` files in the destination folder you select.
- Patient header information is written only into the local Word report and is not transmitted to any external service.
- The application may store minimal configuration data (e.g., last-used folder, window layout, or theme settings) in local files next to the application or in a user-specific configuration directory, depending on implementation.
- The OpenAI API key is requested at runtime for each session and is not saved permanently unless explicitly implemented by the user.

---

## 8) Privacy and Responsibility

AI Clinical Interpreter is designed to support safe use of anonymized clinical data with AI, but it does not replace institutional data-protection procedures.

- All uploaded files must be fully anonymized before being processed with this tool.
- Only anonymized content is sent to the AI provider (e.g., OpenAI).
- Patient header fields remain local and are never transmitted.
- You remain responsible for:
  - Ensuring compliance with data protection laws and institutional policies (e.g., GDPR, HIPAA).
  - Verifying that all documents are genuinely anonymized before use.
  - Critically reviewing the AI-generated report before using it in clinical communication or decision-making.

The tool is provided as a decision-support aid, not a certified medical device, and must not be used as the sole basis for diagnosis or treatment.

---

## 9) License

This software can be released under the GNU General Public License v3.0 (GPLv3) or another open-source license of your choice.

Example (GPLv3):

This software is released under the GNU General Public License v3.0 (GPLv3).

You are free to use, modify, and redistribute the software under the terms of the GPLv3.

This program is provided “AS IS,” without warranty of any kind.

(Adjust this section to match the license you actually choose and include the corresponding LICENSE file.)

---

## 10) Acknowledgments and Donations

I gratefully acknowledge the support of the Gesellschaft der Freunde der Universität Heidelberg e.V. (Society of Friends of Heidelberg University).

This app is free to use. If it helps your work, please consider a donation to support students and early-career researchers:

- Recipient: Gesellschaft der Freunde der Universität Heidelberg e.V.
- Bank: Deutsche Bank Heidelberg
- IBAN: DE22 6727 0003 0049 4005 00
- BIC (SWIFT): DEUTDESM672
- Reference (optional): Davoudi App Donation
