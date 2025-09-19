# Checkbook_Balancing_Program

Balance Personal Checkbook
A pragmatic, auditable tool for reconciling a personal or small‑business checkbook against bank statements. It reads your transactions from a spreadsheet (Google Sheets or Excel), computes running balances, flags discrepancies, and helps you reconcile cleared vs. uncleared items.

> **Note on files:** This project uses a Google Sheets workbook as the source of truth. An `.xlsx` file is provided as the Excel equivalent of that Google Sheet, with the same column layout and formulas where applicable.

---

## Table of Contents

1. [Why this exists](#why-this-exists)
2. [Key features](#key-features)
3. [How it works (high‑level)](#how-it-works-high-level)
4. [Data model (Spreadsheet schema)](#data-model-spreadsheet-schema)
5. [Project structure](#project-structure)
6. [Setup](#setup)

   * [Google Sheets mode](#google-sheets-mode)
   * [Excel mode](#excel-mode)
7. [Usage](#usage)

   * [Adding transactions](#adding-transactions)
   * [Reconciling to a bank statement](#reconciling-to-a-bank-statement)
   * [Importing from bank CSV/OFX](#importing-from-bank-csvofx)
   * [Reports](#reports)
8. [Configuration](#configuration)
9. [Reconciliation logic (under the hood)](#reconciliation-logic-under-the-hood)
10. [Quality, testing, and auditability](#quality-testing-and-auditability)
11. [Performance notes](#performance-notes)
12. [Troubleshooting](#troubleshooting)
13. [FAQ](#faq)
14. [License](#license)

---

## Why this exists

Balancing a checkbook is still the cleanest way to know exactly where your money went, especially when bank websites lump transactions together or delay pending/cleared statuses. This program gives you a single, reliable ledger with:

* A **truth‑table** of transactions you control
* An **explicit cleared/uncleared** state per line
* A **repeatable reconciliation workflow** that matches your ledger to a statement balance
* **Simple diffusion of errors**: mistakes are localized and easy to trace

## Key features

* **Two storage backends**: Google Sheets (primary) and Excel `.xlsx` (equivalent layout for offline/desktop use).
