# Berichtsassistent 2 — C# Excel Add-in

An Excel Add-in with custom user-defined functions (UDFs) for automated real-time data retrieval from a REST API backend.

## What it does
- Custom Excel functions that fetch live data from a REST API within seconds
- Automated report generation and structured data display directly in Excel cells
- Ribbon UI for configuring server connection
- Supports multiple data types: clients, fiscal years, P&L, balance sheet

## Technologies
- **C# / .NET** — core logic and UDF implementation
- **Excel-DNA** — Excel Add-in and UDF registration framework
- **REST API** — async HTTP data retrieval
- **MS SQL Server** — backend database

## Structure
```
Daten/     # Data layer — UDF functions and API calls
Ribbon/    # Excel Ribbon UI and task pane
```

## Background
Built to replace a slow manual reporting process. Data that previously required multiple steps is now available within seconds directly in Excel via custom functions.
