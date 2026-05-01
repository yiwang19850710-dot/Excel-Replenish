# Excel Replenish System (StockPilot Lite)

## 📌 System Positioning

This is an Excel-based lightweight:

- Demand Planning Tool
- Replenishment Decision System
- Stock Health (DSI) Engine

Target users:
- Small e-commerce sellers
- Shopify / Amazon operators
- Users without data science background

Core advantage:
- Excel-based (no learning curve)
- Simple but powerful decision logic

---

# 🧱 SYSTEM ARCHITECTURE

The system is structured into:

1. Master Data (DB tables)
2. Transaction Data (PO, Sales, Inventory)
3. UI Modules (user interaction)
4. Decision Engine (Forecast + DSI + Replenishment)

---

# 📊 DATABASE STRUCTURE

## 1. Products_DB (tblProducts)

Fields:

- Product_ID (Primary Key)
- SKU
- Product_Name
- Variant_Desc
- Category
- Unit_Cost
- Selling_Price
- Opening_Stock
- Current_Stock
- Reorder_Level
- Safety_Days_Override
- Lead_Time_Days   ⭐ (IMPORTANT FOR REPLENISHMENT)
- Active_Status
- Notes
- Created_At
- Updated_At

---

## 2. Suppliers_DB (tblSuppliers)

Fields:

- Supplier_ID
- Supplier_Name
- Contact_Person
- Email
- Phone
- Address
- Country
- Default_Currency   ⭐
- Payment_Terms      ⭐
- Notes
- Active_Status
- Created_At
- Updated_At

---

## 3. Purchase_DB (tblPurchase)

Fields:

- Purchase_Order_No
- External_PO_No  ⭐ (used for grouping external orders)
- Purchase_Date
- Supplier_ID
- Supplier_Name
- Product_ID
- SKU
- Product_Name
- Qty
- Received_Qty
- Remaining_Qty
- Unit_Cost
- Line_Total
- Amount_Paid
- Balance_Due
- Payment_Status (UNPAID / PARTIAL / PAID)
- Status (OPEN / PARTIAL / CLOSED)
- Line_Status
- Notes
- Source
- Created_At
- Updated_At

---

## 4. Inventory_Log

Tracks all inventory movements:

- IN (Receiving)
- OUT (Sales)
- Adjustment

---

## 5. Sales_DB

- Order-level sales data
- Used for demand calculation

---

# 🧩 UI MODULES

## 1. Dashboard

Future role:
- Command Center
- Key KPIs
- Replenishment alerts
- System reset button (planned)

---

## 2. Products_UI

Functions:
- Create Product
- Update Product
- Load Product

Key logic:
- Lead_Time_Days stored at product level

---

## 3. Suppliers_UI

Functions:
- Create Supplier
- Update Supplier
- Load Supplier

---

## 4. Import_UI

Supports:

- Product Import
- Inventory Import
- PO Import
- Supplier Import

Key features:

- Validation before import
- STOP_ON_ERROR
- IMPORT_VALID_ONLY
- Duplicate handling via UserForm

---

## 5. PO_Inquiry_UI

Functions:

- Search PO by SKU / Supplier / PO No
- Show Open / Partial / All PO
- Load selected PO into Purchase_UI

---

## 6. Stock_Health_UI ⭐ CORE MODULE

Displays:

- 90-day horizontal timeline
- Forecast vs Supply
- DSI (Days Sales of Inventory)

Features:

- Multi-SKU view (up to 5 SKUs)
- Manual override support
- Lead time simulation
- Color-coded DSI

---

## 7. Forecast_UI

Handles:

- Forecast logic
- Trend sensitivity
- Seasonality (auto)
- Promo multiplier

---

# 🧠 DECISION ENGINE (CORE)

## 1. Forecast

Current approach:

- Trend-based
- Optional seasonality
- Adjustable sensitivity

Future improvement:

- Better demand smoothing
- Outlier handling
- Sales-weighted forecast

---

## 2. Stock Health (DSI)

Definition:

DSI = Current Stock / Daily Demand

Used for:

- Risk detection
- Overstock / stockout analysis

---

## 3. Replenishment Logic (NEXT STEP)

To be enhanced:

- When to reorder
- How much to reorder
- Based on:
  - Lead Time
  - Demand
  - Safety stock

---

# 🚀 PRODUCT DIRECTION

Transform system into:

### Excel-based Replenishment Decision Engine

Focus on:

1. Simple UX (non-technical users)
2. High clarity decision outputs
3. Fast simulation (what-if analysis)

---

# ⚠️ DESIGN PRINCIPLES

- Do NOT overcomplicate
- Keep Excel-native experience
- Prefer transparency over black-box logic
- Optimize for usability, not theory

---

# 🎯 NEXT OPTIMIZATION TARGET

1. Replenishment Suggestion Engine
2. Forecast improvement
3. Lead Time usage enhancement
4. Dashboard as Command Center
5. Code modularization


Add workbook structure documentation
