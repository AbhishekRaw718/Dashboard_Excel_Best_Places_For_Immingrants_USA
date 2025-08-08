# City Selection Decision Dashboard (Excel + VBA)

An interactive Excel dashboard to compare and rank U.S. cities using multi-criteria decision analysis. The model normalizes inputs, applies configurable weights, and produces dynamic rankings with pivot visuals, slicers, and macros.

> **File:** `Group12 Dashboard.xlsm` (enable macros)

---

## 🔍 What this dashboard does
- Ranks **30+ cities** on **taxes, crime, cost of living, job market, culture, and quality-of-life** factors.
- Uses **normalization + weighted scoring** to compute an **Overall Rating** per city.
- Provides **interactive comparisons** via **pivot tables, slicers, and charts**.
- Supports quick **“Top 5”** surfacing and **scenario testing** with adjustable weights.

---

## 🗂️ Workbook Structure
- **Dataset** – Raw inputs (city metrics).
- **Normalized Dataset** – Min–max scaled fields (0–1).
- **Weighted Scores** – Weights × normalized metrics → Overall Rating.
- **Dashboard** / **City Dashboard** – Interactive visuals and controls.
- **GraphPivotData** / **TaxPivotData** – Pivot sources for charts/tables.
- **Information** – Weight configuration & data dictionary.
- **top 5 Data** – Helper sheet for Top-N results.
- **Macros** – VBA procedures (e.g., refresh/clear).
- **Sheet1/2/3** – Scratch/backup (if present).

> Tip: Start your review at **Dashboard** → use slicers/filters → adjust **weights** in **Information**.

---

## 🧮 Scoring Method (at a glance)
1. **Normalize** each metric to 0–1 (min–max):
   \[
   x_\text{norm}=\frac{x - \min(x)}{\max(x)-\min(x)}
   \]
   > For “lower-is-better” metrics (e.g., crime, tax), the value is **inverted** post-normalization.
2. **Weighting**: Multiply each normalized metric by its **weight** (set in **Information**).
3. **Overall Rating**: Sum of all weighted metrics → **city rank**.

---

## 🧪 How to use
1. **Open** `Group12 Dashboard.xlsm` in **Excel 2019+ (Windows recommended)**.
2. **Enable Macros** when prompted.
3. Go to **Information** → adjust **weights** (e.g., increase weight on cost of living).
4. Go to **Dashboard**:
   - Use **slicers/filters** to compare segments.
   - View **Top 5** cities and detailed metric breakdowns.
5. Press **Refresh** (macro button) if present, or `Data → Refresh All`.

---

## 🛠️ Customize / Extend
- **Add a metric**:  
  - Append to **Dataset**, mirror into **Normalized Dataset** (add normalization formula),
  - Add a **weight** in **Information**, and include it in **Weighted Scores**.
- **Change scoring**: Swap min–max with z-scores or quantile scaling in **Normalized Dataset**.
- **New visuals**: Point charts to **GraphPivotData** or create new pivots.

---

## ✅ Requirements
- **Excel 2019 or later** (Windows).  
  > Mac works for most features but some slicers/VBA may behave differently.
- Macro security set to **Enable VBA** for this file.

---

## 🔐 Macro Notes
- VBA used for **refresh**, **clearing filters**, and **UI helpers**.
- Code is contained in the **Macros** module.  
  > Review before enabling in restricted environments.

---

## 📊 Key KPIs (examples)
- **Overall Rating** (0–100 scaled)
- **Tax Burden**: State + Local + Total
- **Crime Rate per 100k**
- **Cost of Living Index / Housing Proxy**
- **Job Market Score**
- **Annual Cultural Events**
- **Indian Population %** (demographic proxy)

---

## 📷 Screenshots
_Add screenshots or GIFs here once the repo is live._
- `images/dashboard-overview.png`
- `images/city-comparison.png`
- `images/weights-scenario.gif`

---

## 🧭 Roadmap
- Sensitivity analysis chart (tornado) for weights
- Scenario presets (Budget-friendly / Safety-first / Job-market)
- Export ranked shortlist to CSV via VBA

---

## 📄 License
MIT — feel free to fork and adapt with attribution.

---

## 👤 Author
**Abhishek Rawal**  
Engineering Management | Product & Analytics  
LinkedIn: https://www.linkedin.com/in/abhishek-rawal-2510/ • Email: adrawal2510@gmail.com

