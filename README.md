# SAP Delivery Validation & Export Documentation Automation (Excel VBA)

This Excel VBA macro automates the process of validating outbound deliveries in SAP (via VL02N) and generating export documentation (Packing Lists & Proforma / Invoice for VAT) required for customs requests to Kuehne+Nagel (K&N).

---

## Features
- Reads delivery data from an Excel sheet.
- Connects to SAP GUI using scripting automation.
- Validates each delivery against multiple criteria:
  - Picking status (`C` complete)
  - Transportation group (`Z1`, `Z2`, `Z3`)
  - Net vs gross weight consistency
- Navigates SAP VL02N to open and download the ZPAC output (Packing List).
- Checks existing document flow for invoices or proformas before creating new ones.
- Creates the correct billing document in SAP (`ZCS_VF01`) based on destination country:
  - Invoice for VAT (GB)
  - Proforma (all other countries)
- Updates Excel status columns (S & T) with completion mark and document type.

---

## Benefits
- Eliminates repetitive SAP navigation and data entry.
- Reduces human error in delivery checks and document generation.
- Standardizes export documentation preparation for customs requests.

---

## SAP Macro Flow Diagram
![SAP Automation Flow](https://github.com/LauraPrado22/Delivery-data-check-and-Export-Documentation/blob/main/58F58E63-CE25-4BB3-90EF-CBD0836EAF96.jpeg)

---

## Usage
1. Place delivery data in `Sheet1` starting from row 2.
2. Open SAP GUI and log in to your environment.
3. Run the macro `VL02N_Proforma317_FullMacro`.
4. Follow the on-screen prompts for downloading or generating documents.

---

## Requirements
- Microsoft Excel with VBA enabled
- SAP GUI with scripting enabled
- Access to transactions: `VL02N`, `ZCS_VF01`, `VF03`


