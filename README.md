# 🚗 Vahan Excel Processor

A Java-based utility to fetch vehicle details from the VAHAN API and export complete vehicle data into an Excel file.

---

## 📌 Features

- Read vehicle numbers from Excel
- Fetch data using VAHAN API (MastersIndia)
- Parse XML response dynamically
- Export ALL vehicle fields into Excel columns
- Automatically handles new fields (no code change required)
- Vehicle number always shown in first column

---

## 🛠️ Tech Stack

- Java 8+
- Apache POI (Excel handling)
- org.json (JSON parsing)
- DOM Parser (XML parsing)

---

## 📂 Input Format

Excel file: `vehicle_input.xlsx`

| Vehicle No |
|------------|
| HR55N2344  |
| JH02BJ3738 |

---

## 📤 Output Format

Excel file: `vehicle_output.xlsx`

| vehicle_no | rc_owner_name | rc_status | rc_fuel_desc | rc_maker_model | ... |
|------------|--------------|----------|--------------|----------------|-----|

✔ All XML fields automatically included

---

## ⚙️ Setup Instructions

### 1. Clone Repository

```bash
git clone https://github.com/YOUR_USERNAME/vahan-excel-processor.git
