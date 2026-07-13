from __future__ import annotations

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

APP_VERSION = "2026.07.13-v2.1"

st.set_page_config(
    page_title="Amazon FBA Pallet Allocation Tool",
    page_icon="📦",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
PRODUCT_FILE = BASE_DIR / "Product.xlsx"


@st.cache_data(show_spinner=False)
def get_product_status(product_path: str, modified_time: float) -> tuple[int, int]:
    product_df, warnings_df = load_product_master(product_path)
    return len(product_df), len(warnings_df)


st.title("📦 Amazon FBA Pallet Allocation Tool")
st.caption(
    f"版本：{APP_VERSION} ｜ 固定 Product.xlsx，只需上传同格式订单 Excel。"
)

with st.expander("当前分配规则", expanded=False):
    st.markdown(
        f"""
- 托盘体积上限：**{PALLET_LENGTH_IN:.0f} × {PALLET_WIDTH_IN:.0f} × {PALLET_HEIGHT_IN:.0f} = {PALLET_CAPACITY_IN3:,.0f} in³**。
- 产品资料自动转换：**kg → lb，cm → inch**。
- **混合打托**：使用该 ARN 的固定 Pallets 总数；每托至少 2 个不同 SKU；每托体积不得超限；在硬约束内优先让每个 SKU 分配尽量平均，再平衡各托体积。
- **single sku pallet**：每个 SKU 行自己的 Pallets 数量是该 SKU 的固定托数；按“整除 + 余数”平均分配。
- 若固定托数下无法满足规则，程序不会擅自增加或减少托盘，而是在“错误与警告”中说明原因。
        """
    )

if not PRODUCT_FILE.exists():
    st.error(
        "项目目录中找不到固定产品资料文件 Product.xlsx。"
        "请把 Product.xlsx 放在 streamlit_app.py 同一目录。"
    )
    st.stop()

try:
    sku_count, product_warning_count = get_product_status(
        str(PRODUCT_FILE), PRODUCT_FILE.stat().st_mtime
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("固定产品主数据", f"{sku_count:,} 个唯一 SKU")
    c2.metric("托盘容量", f"{PALLET_CAPACITY_IN3:,.0f} in³")
    c3.metric("产品资料警告", product_warning_count)
    c4.metric("应用版本", APP_VERSION)
except Exception as exc:
    st.error(f"Product.xlsx 读取失败：{exc}")
    st.stop()

uploaded_order = st.file_uploader(
    "上传订单详情 Excel",
    type=["xlsx"],
    help="订单格式应与当前 Order.xlsx 一致。Product.xlsx 已固定在 GitHub 项目中。",
)

if uploaded_order is not None:
    st.success(f"已上传：{uploaded_order.name}")

    if st.button("开始计算托盘分配", type="primary", use_container_width=True):
        try:
            with st.spinner("正在检查 UPC/SKU 匹配并计算托盘分配..."):
                # Preflight diagnostic: use exactly the same loader/matcher as allocation.
                uploaded_order.seek(0)
                order_df = load_order(uploaded_order)
                product_df, _ = load_product_master(PRODUCT_FILE)
                diagnostic_df = enrich_order_with_product(order_df, product_df)

                unmatched_df = diagnostic_df.loc[~diagnostic_df["_matched"]].copy()
                matched_count = int(diagnostic_df["_matched"].sum())
                unmatched_count = int((~diagnostic_df["_matched"]).sum())

                d1, d2, d3 = st.columns(3)
                d1.metric("订单有效行", len(diagnostic_df))
                d2.metric("产品资料已匹配", matched_count)
                d3.metric("未匹配", unmatched_count)

                if unmatched_count > 0:
                    st.error(
                        f"有 {unmatched_count} 条订单无法匹配当前 GitHub 中的 Product.xlsx。"
                    )
                    show_cols = [
                        c
                        for c in [
                            "_row_no",
                            ORDER_ARN,
                            ORDER_MODEL,
                            ORDER_UPC,
                            "_upc_key",
                            "_model_key",
                        ]
                        if c in unmatched_df.columns
                    ]
                    display_df = unmatched_df[show_cols].rename(
                        columns={
                            "_row_no": "Excel Row",
                            "_upc_key": "Normalized UPC",
                            "_model_key": "Normalized Model",
                        }
                    )
                    st.subheader("未匹配明细")
                    st.dataframe(display_df, use_container_width=True, hide_index=True)
                    st.info(
                        "如果这里显示的未匹配数量不是 0，请确认 Streamlit 页面顶部版本必须是 "
                        f"{APP_VERSION}，且固定产品主数据应显示 347 个唯一 SKU。"
                    )
                    st.stop()

                uploaded_order.seek(0)
                result = allocate_orders(uploaded_order, PRODUCT_FILE)
                output_bytes = build_output_excel(result)

            ok_count = int((result.arn_summary["Status"] == "OK").sum())
            failed_count = int((result.arn_summary["Status"] != "OK").sum())
            pallet_count = len(result.pallet_summary)

            m1, m2, m3 = st.columns(3)
            m1.metric("成功 ARN", ok_count)
            m2.metric("失败 ARN", failed_count)
            m3.metric("已生成托盘", pallet_count)

            st.subheader("ARN 汇总")
            st.dataframe(result.arn_summary, use_container_width=True, hide_index=True)

            if not result.issues.empty:
                st.subheader("⚠️ 错误与警告")
                st.dataframe(result.issues, use_container_width=True, hide_index=True)

            if not result.pallet_summary.empty:
                st.subheader("托盘汇总")
                st.dataframe(result.pallet_summary, use_container_width=True, hide_index=True)

            if not result.detail.empty:
                with st.expander("查看托盘分配明细"):
                    st.dataframe(result.detail, use_container_width=True, hide_index=True)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="⬇️ 下载托盘分配结果 Excel",
                data=output_bytes,
                file_name=f"Amazon_FBA_Pallet_Allocation_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as exc:
            st.exception(exc)
else:
    st.info("请先上传订单 Excel。")
