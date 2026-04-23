# RPA Certificate Processing Bot

An unattended UiPath RPA process that automates certificate retrieval and download from the PM system by processing incoming Outlook emails, validating request fields, executing web automation via Chrome, and sending result notifications.

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Project Structure](#project-structure)
- [Configuration](#configuration)
- [Execution Flow](#execution-flow)
- [Processing Modes](#processing-modes)
- [Email Validation Rules](#email-validation-rules)
- [Error Handling & Notifications](#error-handling--notifications)
- [Testing](#testing)
- [Deployment](#deployment)

---

## Overview

| Field | Value |
|---|---|
| Project Name | CACertificateClose |
| Description | Update and Download Certificate in PM System |
| Version | 1.0.1 |
| UiPath Studio | 20.10.2+ |
| Target Framework | .NET Framework (Legacy) |
| Expression Language | C# |
| Entry Point | `Main.xaml` |

The bot monitors a designated Outlook mailbox for certificate application emails (出证申请), parses and validates the required fields, then automates the PM system in Chrome to download or update the corresponding certificate. Results are logged and result emails are sent automatically.

---

## Prerequisites

- **UiPath Studio** 20.10.x or later (Legacy project)
- **Microsoft Office** (Outlook + Excel) installed and configured on the robot machine
- **Google Chrome** with UiPath Chrome Extension installed
- **Network access** to the PM system and case link URLs
- **SMTP server** credentials for outbound email notifications
- An Outlook account with access to the TA mailbox being monitored

### Key UiPath Package Dependencies

| Package | Version |
|---|---|
| UiPath.System.Activities | 21.10.2 |
| UiPath.UIAutomation.Activities | 21.10.3 |
| UiPath.Mail.Activities | 1.12.2 |
| UiPath.Excel.Activities | 2.11.4 |
| UiPath.IntelligentOCR.Activities | 4.10.0 |
| UiPath.FTP.Activities | 1.0.6710.16240 |
| UiPath.Credentials.Activities | 1.1.6479.13204 |
| SGS-CSTC.Activities.LocalLog | 1.0.0 |
| SGS-CSTC.Activities.Requests | 4.1.2 |
| SGS-CSTC.Activities.BaiduOCR | 1.0.0 |
| PDFToImageActivities | 1.1.1 |

All custom NuGet packages are included under the `Activity/` directory.

---

## Project Structure

```
rpa-certificate/
├── Main.xaml                          # Entry point — orchestrates the full process
├── project.json                       # UiPath project manifest & dependency list
│
├── Config/
│   ├── Config.xlsx                    # Runtime settings, constants, and credentials
│   └── EmailMappingList.xlsx          # Sender-to-name mapping for email routing
│
├── Framework/                         # Reusable utility workflows
│   ├── InitAllSettings.xaml           # Loads Config.xlsx into a runtime Dictionary
│   ├── GetAppCredentials.xaml         # Retrieves stored application credentials
│   ├── GetOutlookTAEmail.xaml         # Fetches TA emails from Outlook
│   ├── GetOutlookTAEmailResend.xaml   # Fetches emails for resend scenarios
│   ├── GetOutlookProcertEmail.xaml    # Fetches Procert-specific emails
│   ├── SendTemplateMail.xaml          # Sends templated email responses
│   ├── QueueProcess.xaml              # Queue transaction management
│   ├── QueueWriteLock.xaml            # Acquires a write lock on the queue
│   ├── QueueUnLock.xaml               # Releases the queue write lock
│   ├── CreatePath.xaml                # Creates required directory paths
│   ├── CreateTransactionDataTable.xaml# Initializes the transaction data structure
│   ├── InitFullNameMapping.xaml       # Builds the full-name mapping dictionary
│   ├── GetDaysBetweenDate.xaml        # Calculates date differences for retry logic
│   ├── WriteIntoResult.xaml           # Writes processing results to output files
│   ├── TakeScreenshot.xaml            # Captures screenshots for audit/debug
│   ├── DoConfig.xaml                  # Applies configuration values at runtime
│   └── listOfDictToDT.xaml            # Converts a list of dictionaries to DataTable
│
├── Process/                           # Business logic workflows
│   ├── CleanupAndPrep.xaml            # Kills Chrome/Excel before starting
│   ├── CloseAllApplications.xaml      # Gracefully closes open applications
│   ├── KillAllProcesses.xaml          # Force-terminates lingering processes
│   ├── ProcessTransactionSimple.xaml  # Handles simple certificate transactions
│   ├── ProcessTransactionOnline.xaml  # Handles online certificate transactions
│   ├── ProcessTransactionOffline.xaml # Handles offline certificate transactions
│   │
│   └── Certificate/                   # Certificate-specific automation
│       ├── ChromePMSimple.xaml        # PM system Chrome automation (simple flow)
│       ├── ChromePMOnline.xaml        # PM system Chrome automation (online flow)
│       ├── ChromePMOffline.xaml       # PM system Chrome automation (offline flow)
│       ├── HandleProcess.xaml         # Core certificate processing logic
│       ├── FindEmail.xaml             # Locates target emails
│       ├── FindEmailResend.xaml       # Handles resend email lookup
│       ├── PickEmail.xaml             # Selects the email to process
│       ├── GetEmailMapping.xaml       # Resolves email-to-name mapping
│       ├── DownloadFile.xaml          # Downloads certificate files
│       ├── ReOpen.xaml                # Reopens applications if closed unexpectedly
│       ├── SendDailyReport.xaml       # Generates and sends the daily summary report
│       ├── SendSMTPEmail.xaml         # Standard result notification
│       ├── SendSMTPEmailExcept.xaml   # Exception notification
│       ├── SendSMTPEmailFailed.xaml   # Failure notification
│       ├── SendSMTPEmailGCC.xaml      # GCC-specific notification
│       └── SendSMTPEmailPending.xaml  # Pending status notification
│
├── Activity/                          # Local NuGet package cache
│   └── *.nupkg                        # Custom and third-party activity packages
│
└── Test/                              # Test workflows for development use
    ├── TestSequence.xaml
    ├── TestSequence02.xaml
    ├── TestSequenceTemp.xaml
    ├── GetOutlookTAEmailTest.xaml
    ├── PDFToImageTest.xaml
    └── ABCSequence.xaml
```

---

## Configuration

All runtime parameters are managed through Excel files in the `Config/` directory. **Do not hard-code values in workflows.**

### Config.xlsx

Contains three sheets:

| Sheet | Purpose |
|---|---|
| **Settings** | Folder paths, Outlook mailbox names, retry counts, processing flags |
| **Constants** | Static values used across workflows (e.g. email subjects, field names) |
| **Credentials** | Application login credentials (loaded via `GetAppCredentials.xaml`) |

> Credential fields and any field named `Private:*` or containing `password` are excluded from UiPath logs automatically.

### EmailMappingList.xlsx

Maps email sender addresses to display names. Used by `GetEmailMapping.xaml` to determine how to address recipients in outbound notifications.

### Local NuGet Feed

The `Activity/` directory serves as a local NuGet package source. Add it to UiPath Studio's package sources before opening the project:

**Manage Sources → Add → Point to** `<project_root>/Activity/`

---

## Execution Flow

```
Main.xaml
  │
  ├─ 1. Cleanup & Init
  │     CleanupAndPrep       — kill Chrome/Excel processes
  │     InitAllSettings      — load Config.xlsx into Dictionary
  │     GetAppCredentials    — retrieve application credentials
  │
  ├─ 2. Email Retrieval
  │     GetOutlookTAEmail / GetOutlookTAEmailResend
  │     FindEmail / PickEmail
  │
  ├─ 3. Email Validation
  │     Extract: Certificate Type, Import Record No.
  │     Validate: 出证申请, 案件链接, Procert Project No.
  │     Check: case link accessibility (NB-GCC special handling)
  │
  ├─ 4. Certificate Processing
  │     ProcessTransactionSimple  →  ChromePMSimple   (simple flow)
  │     ProcessTransactionOnline  →  ChromePMOnline   (online flow)
  │     ProcessTransactionOffline →  ChromePMOffline  (offline flow)
  │         └─ HandleProcess — core PM system automation
  │             DownloadFile — save certificate files
  │
  ├─ 5. Result Logging & Notification
  │     WriteIntoResult      — write results to log file
  │     SendSMTPEmail*       — send appropriate notification email
  │     SendDailyReport      — daily summary (scheduled run)
  │
  └─ 6. Cleanup
        CloseAllApplications
        KillAllProcesses
```

---

## Processing Modes

The bot supports three certificate processing modes, determined by the email content:

| Mode | Workflow | Description |
|---|---|---|
| **Simple** | `ChromePMSimple.xaml` | Standard certificate with no online/offline distinction |
| **Online** | `ChromePMOnline.xaml` | Certificate issued through an online inspection process |
| **Offline** | `ChromePMOffline.xaml` | Certificate issued through an offline inspection process |

Each mode invokes the corresponding Chrome automation workflow against the PM system.

---

## Email Validation Rules

Incoming emails must pass the following validation checks before processing begins. Failures trigger exception notification emails.

| Check | Field | Notes |
|---|---|---|
| Certificate Type | Email subject | Parsed from subject line |
| Import Record No. | Email subject | Parsed from subject line |
| Procert Project No. | Email body | Must be present and non-empty |
| 出证申请 (Certificate Application) | Email body | Required field |
| 案件链接 (Case Link) | Email body | URL must be present |
| Case Link Accessibility | URL reachability | Special bypass for NB-GCC case type |
| "For CS" Folder | File system | Required folder structure must exist |

---

## Error Handling & Notifications

The bot sends different SMTP email types depending on the outcome:

| Scenario | Workflow |
|---|---|
| Successful processing | `SendSMTPEmail.xaml` |
| Processing failed | `SendSMTPEmailFailed.xaml` |
| Unhandled exception | `SendSMTPEmailExcept.xaml` |
| Pending / awaiting data | `SendSMTPEmailPending.xaml` |
| GCC-specific outcome | `SendSMTPEmailGCC.xaml` |
| Daily summary | `SendDailyReport.xaml` |

Screenshots are captured automatically on exceptions via `TakeScreenshot.xaml` and attached to logs for debugging.

---

## Testing

Test workflows are located in the `Test/` directory and are intended for development use only — they should not be published to production.

| Workflow | Purpose |
|---|---|
| `TestSequence.xaml` | General end-to-end sequence test |
| `TestSequence02.xaml` | Secondary sequence test variant |
| `TestSequenceTemp.xaml` | Temporary/scratch test |
| `GetOutlookTAEmailTest.xaml` | Isolated email retrieval test |
| `PDFToImageTest.xaml` | PDF-to-image conversion test |
| `ABCSequence.xaml` | Basic workflow connectivity test |

To run a test, open the corresponding `.xaml` in UiPath Studio and use **Run File** (not **Run Project**).

---

## Deployment

1. Open the project in **UiPath Studio 20.10.x**.
2. Ensure the local `Activity/` folder is registered as a NuGet package source.
3. Restore packages via **Manage Packages** if prompted.
4. Update `Config/Config.xlsx` with the target environment's settings, credentials, and folder paths.
5. Update `Config/EmailMappingList.xlsx` with the correct sender-to-name mappings.
6. Publish the project to **UiPath Orchestrator** or run directly from Studio for testing.
7. Schedule or trigger the process as required by the business schedule.

> **Note:** The robot machine must have Outlook configured and logged in, Chrome installed with the UiPath extension, and network access to the PM system URLs before execution.
