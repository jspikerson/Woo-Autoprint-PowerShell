# WooCommerce Auto-Print Packing Slips (PowerShell 7)

A PowerShell 7 script that watches your WooCommerce store and prints **packing slips** for new orders. Each slip is rendered as **HTML → PDF** (via Edge/Chrome headless) and sent to your printer.

## Features

- Tracks the **last printed order** (no duplicates)
- Prints selected statuses (default: `processing`)
- Clean table: **Qty | Item | Shipped | BO**
- Shows **Product Add-Ons** from `_pao_ids` under each item (key/value pairs)  
  ↳ URL values (e.g., uploaded images) are replaced with **“Image uploaded by customer”**
- Appends **Customer Notes** (checkout note + any `/orders/{id}/notes` with `customer_note=true`)
- Saves per-order **HTML + PDF** copies in `printed/`

---

## Table of Contents

- [Requirements](#requirements)  
- [Folder Layout](#folder-layout)  
- [Setup](#setup)  
- [Configuration](#configuration)  
- [First Run](#first-run)  
- [Run on a Schedule](#run-on-a-schedule)  
- [Silent Printing (Optional)](#silent-printing-optional)  
- [Customization](#customization)  


---

## Requirements

- **Windows 10/11**
- **PowerShell 7+** (`pwsh`)
- **Microsoft Edge** or **Google Chrome** installed (used headless to create PDFs)
- **WooCommerce REST API** key for an account that can read orders
- WordPress **Permalinks** not set to “Plain”
- A printer available to the Windows account that runs the script

---

## Folder Layout

your-repo/
├─ print-woo.ps1 # the script
├─ README.md
├─ .gitignore
└─ (created at runtime)
├─ woo-cred.xml # encrypted creds (per Windows user)
├─ woo-print-state.json # remembers last printed order id
└─ printed/
├─ order-1234.html
└─ order-1234.pdf


---

## Setup

1. **Clone this repo** (or copy the script into your repo/folder).

2. **Edit the config block** at the top of `print-woo.ps1`:

   '''powershell
   # CONFIG — Edit these
   $Store           = 'https://yourstore.com'  # no trailing slash
   $StatusesToPrint = @('processing')
   $PrinterName     = $null                    # or 'Your Printer Name'
   $LogoPath        = $null                    # optional PNG for your logo
   $PageSize        = 'Letter'                 # or 'A4'

3. Create Woo API keys
WooCommerce → Settings → Advanced → REST API → Add key

Choose a user with access to orders (Shop Manager or Admin)

Permissions: Read (or Read/Write if you plan to extend later)

Copy Consumer key (ck_...) and Consumer secret (cs_...)

## First Run

Use -Backlog once to print existing orders and establish the starting point:

pwsh -File ".\print-woo.ps1" -Backlog


You’ll be prompted to enter your Woo credentials:

Username = Consumer Key (ck_...)

Password = Consumer Secret (cs_...)

The script will create:

woo-cred.xml (encrypted for the current Windows user)

woo-print-state.json (tracks the last printed order)

printed\order-XXXX.(html|pdf) for each order printed

Subsequent runs without -Backlog will print only new orders.

# Run on a Schedule

Use Task Scheduler to run every minute:

1. Create Task…

2. General

Run whether user is logged on or not (use an account that can access the printer)

3. Triggers → New…

Begin the task: On a schedule → Daily

Repeat task every: 1 minute → for a duration of: Indefinitely

4. Actions → New…

Program/script: pwsh.exe

Arguments:

-NoProfile -ExecutionPolicy Bypass -File "C:\Path\to\print-woo.ps1"

Start in: C:\Path\to

5. Save (enter the account password when prompted).

## Silent Printing (Optional)

Windows shell printing can pop a viewer briefly. To avoid that, use SumatraPDF (portable):

Download SumatraPDF.exe and place it next to print-woo.ps1.

Replace the script’s Print-Pdf function with the silent version that calls Sumatra’s CLI:

Flags: -silent, -exit-when-done, -print-to-default or -print-to "Your Printer Name"

Optional -print-settings: paper=letter,scaling=fit,duplex,long-edge

Set $PrinterName in the config (or leave $null to use the default printer).

You can keep both implementations: prefer Sumatra when present, fallback to shell printing otherwise.

## Quick Commands
# First run (print existing orders once)
pwsh -File ".\print-woo.ps1" -Backlog

# Regular run (only new orders)
pwsh -File ".\print-woo.ps1"

# Allow local scripts, if needed
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
