from datetime import datetime
from pathlib import Path
import streamlit as st
from allocator import ORDER_ARN, ORDER_MODEL, ORDER_UPC, allocate_orders, build_output_excel, enrich_order_with_product, load_order, load_product_master
st.set_page_config(page_title='Amazon FBA Pallet Allocation Tool', page_icon='📦', layout='wide')
PRODUCT_FILE = Path(__file__).resolve().parent / 'Product.xlsx'
st.title('📦 Amazon FBA Pallet Allocation Tool')
if not PRODUCT_FILE.exists():
    st.error('Product.xlsx not found.')
    st.stop()
uploaded_order = st.file_uploader('上传订单 Excel', type=['xlsx'])
if uploaded_order is not None:
    if st.button('开始计算托盘分配', type='primary', use_container_width=True):
        try:
            with st.spinner('正在计算...'):
                uploaded_order.seek(0)
                order_df = load_order(uploaded_order)
                product_df, _ = load_product_master(PRODUCT_FILE)
                diagnostic_df = enrich_order_with_product(order_df, product_df)
                unmatched_df = diagnostic_df.loc[~diagnostic_df['_matched']].copy()
                if not unmatched_df.empty:
                    st.error(f'有 {len(unmatched_df)} 条订单无法匹配 Product.xlsx')
                    show_cols = [c for c in [ORDER_ARN, ORDER_UPC, ORDER_MODEL] if c in unmatched_df.columns]
                    st.dataframe(unmatched_df[show_cols].rename(columns={ORDER_ARN: 'ARN', ORDER_UPC: 'UPC', ORDER_MODEL: 'Model Number'}), use_container_width=True, hide_index=True)
                    st.stop()
                uploaded_order.seek(0)
                result = allocate_orders(uploaded_order, PRODUCT_FILE)
                output_bytes = build_output_excel(result)
            ok_count = int((result.arn_summary['Status'] == 'OK').sum())
            failed_count = int((result.arn_summary['Status'] != 'OK').sum())
            pallet_count = len(result.pallet_summary)
            c1, c2, c3 = st.columns(3)
            c1.metric('成功 ARN', ok_count)
            c2.metric('失败 ARN', failed_count)
            c3.metric('托盘数', pallet_count)
            if not result.issues.empty:
                st.error('部分 ARN 分配失败')
                st.dataframe(result.issues, use_container_width=True, hide_index=True)
            detail = result.detail[['Shipment ID (ARN)', 'Pallet No.', 'UPC/SKU', 'Model Number', 'Allocated Boxes']].copy()
            detail = detail.rename(columns={'Shipment ID (ARN)': 'ARN', 'Pallet No.': 'Pallet', 'UPC/SKU': 'UPC', 'Allocated Boxes': 'Quantity'})
            if not detail.empty:
                detail = detail.sort_values(['ARN', 'Pallet', 'Model Number'], kind='stable').reset_index(drop=True)
            st.dataframe(detail, use_container_width=True, hide_index=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            st.download_button('下载结果 Excel', data=output_bytes, file_name=f'Amazon_FBA_Pallet_Allocation_{timestamp}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary', use_container_width=True)
        except Exception as exc:
            st.error(str(exc))
