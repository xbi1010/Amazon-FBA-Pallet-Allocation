from datetime import datetime
from pathlib import Path
import streamlit as st
from allocator import (
    ORDER_ARN,
    ORDER_MODEL,
    ORDER_UPC,
    PALLET_CAPACITY_IN3,
    PALLET_HEIGHT_IN,
    PALLET_LENGTH_IN,
    PALLET_WIDTH_IN,
    allocate_orders,
    build_output_excel,
    enrich_order_with_product,
    load_order,
    load_product_master,
)

st.set_page_config(
    page_title="Amazon FBA Pallet Allocation Tool",
    page_icon="📦",
    layout="wide",
)

PRODUCT_FILE = Path(__file__).resolve().parent / "Product.xlsx"

st.title("📦 Amazon FBA Pallet Allocation Tool")
st.caption("Amazon FBA 托盘分配工具")

with st.expander("当前分配规则 / Current Allocation Rules", expanded=False):
    st.markdown(
        f'''
**托盘容量 / Pallet Capacity**

- 最大尺寸 / Maximum dimensions: **{PALLET_LENGTH_IN:.0f} × {PALLET_WIDTH_IN:.0f} × {PALLET_HEIGHT_IN:.0f} in**
- 最大计算体积 / Maximum calculated volume: **{PALLET_CAPACITY_IN3:,.0f} in³**

**混合打托 / Mixed Pallet**

- 使用同一 ARN 对应的固定托盘总数，不自动增加或减少托盘。  
  Use the fixed pallet count assigned to the same ARN. The system will not automatically add or remove pallets.
- 每个托盘至少包含 2 个不同 SKU。  
  Each pallet must contain at least 2 different SKUs.
- 每个托盘的计算总体积不得超过托盘容量。  
  The calculated total volume of each pallet must not exceed the pallet capacity.
- 在满足以上条件的前提下，尽量平均分配各 SKU，并平衡各托盘体积。  
  Within these constraints, SKU quantities are distributed as evenly as possible while balancing pallet volume.

**单一 SKU 打托 / Single SKU Pallet**

- 每个 SKU 使用订单中该行指定的固定 Pallets 数量。  
  Each SKU uses the fixed pallet quantity specified on its order line.
- 数量按“平均分配 + 余数自动分布”处理。  
  Quantity is allocated using even distribution with automatic remainder distribution.
- 例如 / Example: **276 ÷ 10 = 6 pallets × 28 + 4 pallets × 27**
- 每个托盘的计算总体积不得超过托盘容量。  
  The calculated total volume of each pallet must not exceed the pallet capacity.

**异常处理 / Exceptions**

- 如果固定托盘数无法满足规则，系统会显示错误，不会擅自修改托盘数量。  
  If the fixed pallet count cannot satisfy the rules, the system reports an error instead of changing the pallet quantity.
'''
    )

if not PRODUCT_FILE.exists():
    st.error(
        "找不到 Product.xlsx。请将 Product.xlsx 放在 streamlit_app.py 同一目录。 "
        "/ Product.xlsx not found. Place it in the same directory as streamlit_app.py."
    )
    st.stop()

uploaded_order = st.file_uploader(
    "上传订单 Excel / Upload Order Excel",
    type=["xlsx"],
)

if uploaded_order is not None:
    st.success(f"已上传 / Uploaded: {uploaded_order.name}")

    if st.button(
        "开始计算托盘分配 / Calculate Pallet Allocation",
        type="primary",
        use_container_width=True,
    ):
        try:
            with st.spinner("正在计算 / Calculating..."):
                uploaded_order.seek(0)
                order_df = load_order(uploaded_order)
                product_df, _ = load_product_master(PRODUCT_FILE)
                diagnostic_df = enrich_order_with_product(order_df, product_df)
                unmatched_df = diagnostic_df.loc[~diagnostic_df["_matched"]].copy()

                if not unmatched_df.empty:
                    st.error(
                        f"有 {len(unmatched_df)} 条订单无法匹配 Product.xlsx "
                        f"/ {len(unmatched_df)} order lines could not be matched to Product.xlsx"
                    )
                    show_cols = [
                        c
                        for c in [ORDER_ARN, ORDER_UPC, ORDER_MODEL]
                        if c in unmatched_df.columns
                    ]
                    st.dataframe(
                        unmatched_df[show_cols].rename(
                            columns={
                                ORDER_ARN: "ARN",
                                ORDER_UPC: "UPC",
                                ORDER_MODEL: "Model Number",
                            }
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )
                    st.stop()

                uploaded_order.seek(0)
                result = allocate_orders(uploaded_order, PRODUCT_FILE)
                output_bytes = build_output_excel(result)

            ok_count = int((result.arn_summary["Status"] == "OK").sum())
            failed_count = int((result.arn_summary["Status"] != "OK").sum())
            pallet_count = len(result.pallet_summary)

            c1, c2, c3 = st.columns(3)
            c1.metric("成功 ARN / Successful ARNs", ok_count)
            c2.metric("失败 ARN / Failed ARNs", failed_count)
            c3.metric("托盘数 / Pallets", pallet_count)

            if not result.issues.empty:
                st.error("部分 ARN 分配失败 / Some ARNs failed allocation")
                st.dataframe(
                    result.issues,
                    use_container_width=True,
                    hide_index=True,
                )

            detail = result.detail[
                [
                    "Shipment ID (ARN)",
                    "Pallet No.",
                    "UPC/SKU",
                    "Model Number",
                    "Allocated Boxes",
                ]
            ].copy()

            detail = detail.rename(
                columns={
                    "Shipment ID (ARN)": "ARN",
                    "Pallet No.": "Pallet",
                    "UPC/SKU": "UPC",
                    "Allocated Boxes": "Quantity",
                }
            )

            if not detail.empty:
                detail = detail.sort_values(
                    ["ARN", "Pallet", "Model Number"],
                    kind="stable",
                ).reset_index(drop=True)

            st.subheader("托盘分配结果 / Pallet Allocation Result")
            st.dataframe(
                detail,
                use_container_width=True,
                hide_index=True,
            )

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "下载结果 Excel / Download Result Excel",
                data=output_bytes,
                file_name=f"Amazon_FBA_Pallet_Allocation_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as exc:
            st.error(str(exc))
