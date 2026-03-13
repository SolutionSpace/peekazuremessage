# PeekAzureMessage

PeekAzureMessage is a utility designed to read, archive, and extract messages from an Azure Service Bus Queue to local Excel and text files. It operates in two primary modes: reading messages from the queue (based on date ranges) and extracting specific messages across previously generated logs based on search terms.

## Table of Contents
- [Features](#features)
- [Usage Options](#usage-options)
  - [1. Read Messages](#1-read-messages)
  - [2. Extract Messages](#2-extract-messages)

## Features

### 1. Service Bus Queue Peeking (`ReadMessages.cs`)
- Peeks messages from an Azure Service Bus Queue in batches (default: 500) without consuming or deleting them from the queue.
- Support for filtering messages by **Date Range** (`FROM date` and `TO date`).
- Optional filtering starting from a specific **Sequence Number**.
- Auto-stops reading if limits are exceeded or ranges are entirely covered.
- **Excel Export**: Exports essential message metadata (Sequence Number, Message ID, Enqueued Time, State, Subject) to an Excel Spreadsheet using `ClosedXML`.
- **Large Message Handling**: For message bodies exceeding 32,767 characters (Excel's cell character limit), the software accurately identifies this, truncates the preview in the spreadsheet, and saves the full, unabridged message content in isolated `.txt` files.

### 2. Message Extraction & Merging (`ExtractMessages.cs`)
- Prompts the user to search using **comma-separated search terms**.
- Scans through the "Message Text" column of all generated `.xlsx` log files located in the output folder.
- Case-insensitive search inside the message bodies.
- Consolidates any matching occurrences by aggregating all rows that share the same **Message ID** and produces a unified search-results Excel file (`SearchResults_{timestamp}.xlsx`) containing matching entries alongside their original contexts.


## Usage Options

When executing the application, a prompt in the console will ask you for a choice:

```
Enter your choice  ( 1 to Read Messages and  2 to Extract Messages ):
```

### 1. Read Messages
By choosing option **1**, the program performs an analysis of an Azure queue.
- Prompts you for a `FROM date (yyyy-MM-dd)`.
- Prompts you for a `TO date` (defaults to the end of the current day).
- Prompts for an optional `starting Sequence Number` (defaults to 0).
- Results generate an Excel spreadsheet formatted as `Messages_YYYY-MM-DD_to_YYYY-MM-DD.xlsx` loaded with logs.
- Extensively large message contents export cleanly into a `FullMessages` folder as text items.

### 2. Extract Messages
By choosing option **2**, the program processes local generated files to find specific strings.
- Prompts you to enter your search terms separated by commas (e.g. `Error, exception, payloadKey`).
- Sequentially scans the existing generated Spreadsheets inside the `ExcelSheets` directory. 
- Combines the related message items sharing matched Message IDs and drops a consolidated Spreadsheet file sequentially named inside your local `SearchResults` directory.

