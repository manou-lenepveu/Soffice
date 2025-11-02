![.NET](https://img.shields.io/badge/.NET-8.0-purple)
![LibreOffice](https://img.shields.io/badge/requires-LibreOffice-orange)

# Soffice: ODS ↔ XLSX Converter

.NET CLI tool using **LibreOffice** to convert spreadsheet files between **ODS** (LibreOffice) and **XLSX** (Excel), preserving formulas, styles, and multiple sheets.

---

## Prerequisites

| Dependency       | Installation (Ubuntu/Debian)           |
|------------------|----------------------------------------|
| .NET 8 SDK       | `sudo apt install dotnet-sdk-8.0`      |
| **LibreOffice**  | `sudo apt install libreoffice`         |

> Check: `soffice --version`

---

## Installation

```bash
git clone https://github.com/manou-lenepveu/Soffice.git
cd Soffice
dotnet restore

