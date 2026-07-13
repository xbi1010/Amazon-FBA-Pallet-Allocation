# Amazon FBA Pallet Allocation Tool

A Streamlit app for allocating Amazon FBA orders to fixed pallet counts.

## Files

- `streamlit_app.py` — Streamlit web UI
- `allocator.py` — allocation and validation logic
- `Product.xlsx` — fixed product master data
- `requirements.txt` — Python dependencies

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Business rules

- Pallet capacity: `48 × 40 × 72 = 138,240 in³`
- Product data: kg -> lb, cm -> inch
- Mixed pallet:
  - fixed ARN pallet total
  - at least 2 different SKUs per pallet
  - pallet volume <= 138,240 in³
  - SKU quantities are distributed as evenly as possible, then pallet volumes are balanced
- Single SKU pallet:
  - each SKU row has its own fixed Pallets count
  - quantity is split using quotient + remainder
  - example: 276 / 10 => 6 pallets of 28 + 4 pallets of 27
- Fixed pallet counts are never silently changed. Infeasible ARNs are reported as errors.
