# The Ledger of the Goldenthrone

**An Excel-based merchant economy and campaign management tool for Dungeons & Dragons 5e.**

The Ledger of the Goldenthrone started as a solution to a problem at my own table: shopping sessions involved constantly jumping between books, checking prices, tracking what each player was buying, and deciding what a merchant actually had available.

What began as a merchant inventory generator gradually grew into a larger campaign management system.

The basic idea is simple:

> **The players may have all the gold in the world. That doesn't mean the merchant has everything they want to buy.**

The Ledger gives merchants persistent, limited inventories and their own finances while providing the DM with tools for managing shopping, selling, auctions, campaign time, treasure, weather, encounters, and the merchant economy.

The project was originally built for **Tomb of Annihilation** and Port Nyanzaru, so the included merchants, locations, and some campaign systems reflect that adventure.

---

## Features

### Persistent Merchant Inventories

Merchants maintain actual inventories rather than generating an unlimited list of available items every time the players visit.

Items purchased by the party are removed from stock and remain unavailable until they are replenished through the merchant economy.

Different merchants can have their own:

* Inventory capacity
* Item availability
* Item rarity distribution
* Category and subcategory preferences
* Pricing
* Available gold
* Restock schedules
* Persistent stock

This allows individual merchants to feel different from one another while preventing every shop from effectively becoming an unlimited catalog.

### Merchant Economy

Merchants maintain their own available gold as well as their inventory.

Purchases made by the party add gold to the merchant's available funds, while items sold to the merchant reduce those funds.

A merchant therefore may not always have enough gold available to purchase everything the party wants to sell.

### Progressive Restocking

Merchant inventories are tied to campaign time.

As days advance, shops become eligible for restocking according to their individual schedules. Depleted inventory can be replenished and new stock introduced without completely regenerating a merchant every time the party returns.

A Restock Report provides a summary of the changes made during the most recent merchant-economy processing cycle.

### Buying & Selling

The Active Merchant interfaces provide a central location for running shopping sessions.

Items can be selected for purchase or sale, with support for both listed and negotiated pricing.

Completed transactions update the appropriate merchant inventory and finances.

### Item Reservations

Items can be temporarily reserved for the party.

Reserved inventory is protected from normal inventory processing until the reservation expires, allowing a merchant to hold an item for the characters without permanently removing it from the economy.

### Auction System

The Ledger includes an auction system for handling valuable, magical, or unusual items.

Auctions can contain both player-submitted items and randomly generated lots.

Auction values include randomized market behavior so the final price of an item is not always predictable from its base value.

### Calendar of Harptos

An integrated **Calendar of Harptos** tracks campaign time.

The calendar supports normal tendays and Forgotten Realms festival days while providing controls for advancing the campaign.

Campaign time is also connected to other systems in the workbook, including merchant restocking.

### Weather & Encounters

Weather and encounter generation are integrated with the campaign calendar.

This allows the DM to advance campaign time and generate supporting daily information from the same interface.

### Treasure

The workbook contains supporting treasure systems for handling items such as gems and art objects alongside the larger merchant and auction economy.

---

## Requirements

The Ledger is designed primarily for:

**Microsoft Excel 365 for Windows**

The workbook relies extensively on:

* VBA macros
* Excel tables
* Dynamic array formulas
* Modern Excel functions
* Form controls and/or ActiveX controls

**Macros must be enabled for the workbook to function correctly.**

The Ledger is **not designed for Google Sheets or Excel Online**.

Compatibility with older desktop versions of Microsoft Excel and Excel for macOS has not been extensively tested.

---

## Fonts

The Ledger of the Goldenthrone uses custom typography extensively as part of its visual design.

### Modesto

The workbook uses the **Modesto** typeface for many of its themed headings and interface elements.

Modesto is a commercially licensed typeface and **is not distributed with this project**.

Users with access to the appropriate font through **Adobe Fonts** can activate it through Adobe Creative Cloud.

If Modesto is not installed or activated, Microsoft Excel may substitute another available font. The workbook should remain functional, but headings and other interface elements may appear differently or may not fit their intended layout.

### Georgia

The workbook also makes extensive use of **Georgia**, a Microsoft typeface commonly available on Windows systems and with Microsoft Office installations.

Most Windows users should already have Georgia available.

### Font Licensing

**No font files are included or distributed with The Ledger of the Goldenthrone.**

All fonts remain the property of their respective copyright holders and are subject to their respective licensing terms.

Users are responsible for obtaining any fonts necessary to reproduce the intended appearance of the workbook.

---

## Download & Installation

1. Open the **Releases** section of this repository.
2. Download the latest version of **The Ledger of the Goldenthrone**.
3. Extract the downloaded ZIP file.
4. Locate the `.xlsm` workbook.
5. Before opening the workbook, Windows may require the downloaded file to be unblocked.
6. If necessary, right-click the `.xlsm` file and select **Properties**.
7. Check **Unblock** if the option appears, then click **Apply**.
8. Open the workbook in Microsoft Excel.
9. Enable macros/content if prompted.

Because The Ledger contains VBA code, Microsoft Excel and Windows may display security warnings for files downloaded from the internet.

Only enable macros in files downloaded from a source you trust.

---

## Getting Started

The Ledger is divided into several primary interfaces and supporting data sheets.

### Title Page

The main entry point for the workbook.

The Title Page provides navigation to the major sections of The Ledger and displays current campaign information.

### Calendar

Tracks campaign time using the Calendar of Harptos.

The Calendar provides controls for advancing campaign time and connects with systems such as weather, encounters, and merchant restocking.

### Active Merchants

The primary interfaces for interacting with merchants.

These sheets are used to:

* View merchant inventory
* Select items for purchase
* Sell items to merchants
* Negotiate prices
* Reserve inventory
* Complete transactions

### Auction

Manages player-submitted and generated auction lots.

### Items

Contains the master item database used by the merchant, inventory, pricing, and auction systems.

### Treasure

Contains supporting treasure information used throughout The Ledger.

### Shop Inventory

Stores the persistent inventory state of generated merchants.

This is one of the primary data-storage areas used by the merchant economy and generally should not need to be manually edited during normal play.

### Restock Report

Displays the results of the most recent merchant economy/restocking cycle.

---

## Persistent Campaign Data

The Ledger is designed to maintain campaign state between sessions.

Information such as merchant inventory, merchant finances, reservations, and campaign progression may persist inside your workbook.

For that reason:

> **Treat your campaign workbook as a save file.**

Regular backups are strongly recommended.

When updating to a newer release of The Ledger, keep a backup of your existing campaign workbook until you have confirmed that your campaign data has been transferred or preserved correctly.

Do not assume that replacing an existing workbook with a newer release will automatically preserve all campaign state.

---

## Current Version

**v3.7.1**

This is the first publicly distributed version of The Ledger of the Goldenthrone.

The project was originally developed as a tool for my own campaign rather than as a public application. As a result, there may be assumptions, edge cases, usability issues, or bugs that have not surfaced during normal play.

Feedback from other DMs using the workbook is welcome.

---

## Project Status

The Ledger of the Goldenthrone is **actively being developed**.

The current version remains closely tied to the campaign for which it was originally created.

Future development may include:

* Additional merchant-economy features
* Improved inventory management
* Additional automation
* User-interface improvements
* Easier campaign configuration
* Improved update/migration tools
* Greater support for campaigns outside Tomb of Annihilation

Features, workbook structure, tables, formulas, and VBA procedures may change between releases while development continues.

---

## Tomb of Annihilation

The Ledger of the Goldenthrone was originally designed around a **Tomb of Annihilation** campaign.

As a result, the current release contains merchants, locations, economic assumptions, and campaign systems designed with Port Nyanzaru, Chult, and the Forgotten Realms in mind.

The Ledger is therefore **not currently intended to be a completely campaign-neutral D&D economy system**.

Many of its underlying systems could be adapted to other campaigns, and making that process easier is a potential direction for future development.

---

## Updating

New releases will be published through the GitHub **Releases** section using version numbers such as:

`v3.7.1`

Because the workbook contains persistent campaign data, updating is not necessarily as simple as replacing the old workbook.

Before installing a newer version:

1. Back up your current campaign workbook.
2. Read the release notes for the new version.
3. Check for any migration or compatibility instructions.
4. Keep your previous version until you have confirmed that the new version is working correctly.

---

## Feedback & Bug Reports

Feedback is welcome, particularly from DMs who are actually using The Ledger during play.

If you encounter a bug, please submit it through the repository's **Issues** section.

When reporting a problem, please include as much of the following information as possible:

* The Ledger version
* Your version of Microsoft Excel
* What you were attempting to do
* What you expected to happen
* What happened instead
* Any Excel or VBA error message displayed
* A screenshot, if applicable

Please avoid uploading campaign files containing information you do not want to make public.

---

## Suggestions & Feature Requests

Suggestions are welcome.

The Ledger began as a personal campaign tool, so seeing how other DMs use it will help identify which systems would benefit most from additional configuration or automation.

Feature requests can be submitted through GitHub Issues.

Not every suggestion will necessarily be implemented, particularly while the project remains closely tied to its original campaign.

---

## Known Limitations

The Ledger is currently a macro-enabled Microsoft Excel workbook rather than a standalone application.

Some important limitations include:

* Microsoft Excel 365 for Windows is the primary supported environment.
* VBA macros are required.
* Some interface elements may depend on Windows-specific Excel functionality.
* Custom fonts affect the intended visual appearance.
* The current data and configuration are heavily influenced by Tomb of Annihilation.
* Updating between versions may require manual migration of persistent campaign data.
* Compatibility with Excel for macOS has not been extensively tested.
* Google Sheets and Excel Online are not supported.

---

## Disclaimer

The Ledger of the Goldenthrone is an **unofficial, fan-made tabletop roleplaying game tool**.

Dungeons & Dragons, Forgotten Realms, Tomb of Annihilation, and related names, settings, characters, and materials are the property of their respective copyright holders.

This project is not affiliated with, endorsed by, sponsored by, or approved by Wizards of the Coast.

The Ledger of the Goldenthrone is provided as a community-created tool for use by tabletop roleplaying game players and Dungeon Masters.

---

## Third-Party Intellectual Property

References to third-party game systems, settings, adventures, characters, locations, items, fonts, trademarks, or other intellectual property remain the property of their respective owners.

No ownership of third-party intellectual property is claimed.

Third-party fonts referenced by the workbook are **not distributed with this project**.

---

## License

The Ledger of the Goldenthrone and its original code, formulas, workbook design, and original project assets remain subject to the licensing terms provided with this repository.

Third-party intellectual property is excluded from those terms and remains subject to the rights and licenses of its respective owners.

Until explicit redistribution terms are provided, please do not redistribute modified versions of The Ledger of the Goldenthrone as your own release.

---

## Credits

**The Ledger of the Goldenthrone**
Created and developed as a Dungeon Master campaign-management project.

Built with **Microsoft Excel and VBA**.

Originally developed for use with **Tomb of Annihilation**.

Typography includes **Modesto** and **Georgia**. Font files are not distributed with this project.

---

## Support the Project

The Ledger is currently a personal project and is provided for other Dungeon Masters who may find it useful.

If you use it in your campaign, feedback, bug reports, and stories about how it worked at your table are appreciated.
