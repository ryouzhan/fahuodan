import io
from datetime import datetime
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ===================== 页面全局配置 =====================
st.set_page_config(
    page_title="发货单处理工具web版 v1.1",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ===================== 简约高级页面样式 (CSS) =====================
st.markdown("""
<style>
    /* 全局背景与字体 */
    .stApp {
        background-color: #F8FAFC;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        color: #1E293B;
    }

    /* 隐藏 Streamlit 默认顶部条与水印 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* 页面顶部 Header 样式 */
    .header-box {
        padding: 2.2rem 0 1.2rem 0;
        margin-bottom: 1.5rem;
    }
    .header-badge {
        display: inline-block;
        padding: 4px 12px;
        font-size: 0.75rem;
        font-weight: 600;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        background: #EEF2F6;
        color: #475569;
        border-radius: 9999px;
        margin-bottom: 0.6rem;
    }
    .header-title {
        font-size: 1.95rem;
        font-weight: 700;
        color: #0F172A;
        letter-spacing: -0.025em;
        margin: 0;
    }
    .header-subtitle {
        font-size: 0.95rem;
        color: #64748B;
        margin-top: 0.4rem;
        font-weight: 400;
    }

    /* 上传区域精致化卡片 */
    [data-testid="stFileUploader"] {
        background: #FFFFFF;
        border: 1px dashed #CBD5E1;
        border-radius: 12px;
        padding: 1rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.02);
        transition: all 0.2s ease;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #3B82F6;
        box-shadow: 0 4px 12px 0 rgba(59, 130, 246, 0.06);
    }

    /* 指标卡片网格与设计 */
    .metric-container {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1.2rem;
        margin: 1.5rem 0 1.5rem 0;
    }
    .metric-card {
        background: #FFFFFF;
        padding: 1.4rem 1.2rem;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03), 0 4px 12px rgba(0, 0, 0, 0.02);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.06);
    }
    .metric-title {
        font-size: 0.82rem;
        font-weight: 500;
        color: #64748B;
        margin-bottom: 0.4rem;
        text-transform: uppercase;
        letter-spacing: 0.025em;
    }
    .metric-num {
        font-size: 1.75rem;
        font-weight: 700;
        color: #0F172A;
        letter-spacing: -0.03em;
        line-height: 1.2;
    }
    .metric-unit {
        font-size: 0.85rem;
        font-weight: 500;
        color: #94A3B8;
        margin-left: 0.2rem;
    }

    /* 折叠面板 (Expander) 精致化卡片样式 */
    [data-testid="stExpander"] {
        background: #FFFFFF;
        border: 1px solid #E2E8F0 !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
        margin: 1rem 0;
    }

    /* 选项卡 Tabs 简约设计 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 1px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
        font-weight: 500;
        color: #64748B;
        border-radius: 6px 6px 0 0;
    }
    .stTabs [aria-selected="true"] {
        color: #0F172A !important;
        border-bottom: 2px solid #0F172A !important;
        font-weight: 600;
    }

    /* 下载按钮高级样式 */
    .stDownloadButton button {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        font-weight: 500 !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.6rem !important;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.12) !important;
        transition: all 0.18s ease-in-out !important;
    }
    .stDownloadButton button:hover {
        background-color: #1E293B !important;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.22) !important;
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

# ===================== 原始数据处理与表格逻辑（保留原汁原味） =====================
def beautify_excel(file_stream, col_width=15, row_height=20):
    """美化Excel表格，并根据<货件编号>设置整行单元格颜色，处理所有相关Sheet（完全还原原始代码）"""
    try:
        wb = load_workbook(file_stream)
        
        font = Font(name='Calibri', size=11, bold=True)
        alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        colors = [
            "FFCCCC", "CCFFCC", "CCCCFF", "FFFFCC", "CCFFFF", "FFCCFF",
            "FF9999", "99FF99", "9999FF", "FFCC99", "99FFCC", "CCCC99"
        ]

        for sheet_name in ['详细数据', '汇总结果']:
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]
            for col in ws.columns:
                column_letter = col[0].column_letter
                ws.column_dimensions[column_letter].width = col_width

            for row in ws.iter_rows():
                if not row[0].row:
                    continue
                ws.row_dimensions[row[0].row].height = row_height
                for cell in row:
                    if cell.value is None:
                        continue
                    cell.alignment = alignment
                    cell.border = border
                    if cell.row == 1:
                        cell.font = font

            if sheet_name == '详细数据':
                header_row = [cell.value for cell in ws[1]]
                if "货件编号" not in header_row:
                    continue
                item_code_col_idx = header_row.index("货件编号") + 1

                item_codes = {}
                for row in ws.iter_rows(min_row=2, min_col=item_code_col_idx, max_col=item_code_col_idx):
                    val = row[0].value
                    if val and val not in item_codes:
                        item_codes[val] = None

                unique_codes = list(item_codes.keys())
                for idx, code in enumerate(unique_codes):
                    item_codes[code] = colors[idx % len(colors)]

                for row in ws.iter_rows(min_row=2):
                    code_cell = row[item_code_col_idx - 1]
                    code_val = code_cell.value
                    if code_val in item_codes:
                        fill = PatternFill(start_color=item_codes[code_val],
                                           end_color=item_codes[code_val],
                                           fill_type="solid")
                        for cell in row:
                            cell.fill = fill

        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        return output_stream

    except Exception as e:
        raise Exception(f"美化出错: {str(e)}")

def load_and_merge_data(uploaded_file):
    """加载 发货单详情 + 装箱信息 并合并（完全还原原始代码逻辑）"""
    try:
        xl_file = pd.ExcelFile(uploaded_file)
        sheet_names = xl_file.sheet_names

        required_sheets = ['发货单详情', '装箱信息']
        for sheet in required_sheets:
            if sheet not in sheet_names:
                raise KeyError(f"文件缺少必须工作表：{sheet}")

        df_detail = pd.read_excel(uploaded_file, sheet_name='发货单详情')
        df_pack = pd.read_excel(uploaded_file, sheet_name='装箱信息')

        df_detail.rename(columns={
            '货件编号(发货商品)': '货件编号',
            '发货单号(基本信息)': '发货单号',
            'SKU(发货商品)': 'SKU',
            '发货量(发货商品)': '发货量',
            '品名(发货商品)': '品名',
            '商品图片(发货商品)': '商品图片',
            '商品品牌': '供应商'
        }, inplace=True)

        df_pack.rename(columns={
            '货件编号(装箱信息)': '货件编号',
            '发货单号(装箱信息)': '发货单号',
            'SKU(装箱信息)': 'SKU',
            '发货量(装箱信息)': '发货量',
            'ReferenceId(装箱信息)': 'ReferenceId'
        }, inplace=True)

        merged = pd.merge(
            df_detail, df_pack,
            on=['货件编号', '发货单号', 'SKU', '发货量'],
            how='left'
        )

        keep_cols = [
            'SKU', '品名', '商品图片', '发货量', '箱数', '箱规单箱数量', '物流中心编码',
            '供应商', '货件编号', 'ReferenceId', '箱号', '总箱数',
            '外箱重量(kg)', '外箱总重量(kg)', '外箱长(cm)', '外箱宽(cm)', '外箱高(cm)',
            '外箱体积(m³)', '外箱总体积(m³)', '外箱总体积重(kg)', '创建时间', '发货时间', '物流商'
        ]
        exist_cols = [c for c in keep_cols if c in merged.columns]
        merged = merged[exist_cols]

        merged.rename(columns={
            '箱规单箱数量': '单箱数量',
            '总箱数': '总箱数编号'
        }, inplace=True)

        return merged
    except Exception as e:
        raise Exception(f"数据加载合并失败：{str(e)}")

def perform_summary(df):
    """按 物流中心编码 + 货件编号 + ReferenceId 汇总"""
    df_copy = df.copy()
    df_copy['物流中心编码'] = df_copy['物流中心编码'].fillna('未知地址')
    df_copy['货件编号'] = df_copy['货件编号'].fillna('未知编号')
    df_copy['ReferenceId'] = df_copy['ReferenceId'].fillna('无ReferenceId')

    summary = df_copy.groupby(['物流中心编码', '货件编号', 'ReferenceId'], dropna=False).agg(
        品名=('品名', 'first'),
        供应商=('供应商', 'first'),
        总箱数=('箱数', 'sum'),
        外箱总重量=('外箱总重量(kg)', 'sum'),
        外箱总体积=('外箱总体积(m³)', 'sum')
    ).reset_index()

    summary['外箱总体积'] = summary['外箱总体积'].fillna(0)
    summary['外箱总体积重(kg)'] = (summary['外箱总体积'] * 167).round(2)

    return summary

def save_excel(df, summary_df):
    """保存为两个Sheet：详细数据 + 汇总结果"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='详细数据', index=False)
        summary_df.to_excel(writer, sheet_name='汇总结果', index=False)
    output.seek(0)
    return output

# ===================== Web 页面渲染 =====================
st.markdown("""
<div class="header-box">
    <div class="header-badge">Shipment Assistant</div>
    <h1 class="header-title">发货单处理</h1>
    <p class="header-subtitle">by Ryou</p>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "上传发货单 Excel 文件",
    type=["xlsx", "xls"],
    help="支持直接拖拽上传 .xlsx 或 .xls 格式文件"
)

if uploaded_file is not None:
    try:
        with st.spinner("正在处理数据并套用格式..."):
            merged_df = load_and_merge_data(uploaded_file)
            summary_df = perform_summary(merged_df)
            raw_excel = save_excel(merged_df, summary_df)
            excel_bytes = beautify_excel(raw_excel)

        # 1. 核心 KPI 指标卡片
        total_boxes = int(summary_df['总箱数'].sum())
        total_weight = f"{summary_df['外箱总重量'].sum():,.2f}"
        total_vol = f"{summary_df['外箱总体积'].sum():,.2f}"
        total_vol_weight = f"{summary_df['外箱总体积重(kg)'].sum():,.2f}"

        st.markdown(f"""
        <div class="metric-container">
            <div class="metric-card">
                <div class="metric-title">总装箱量</div>
                <div class="metric-num">{total_boxes:,}<span class="metric-unit">箱</span></div>
            </div>
            <div class="metric-card">
                <div class="metric-title">实重合计</div>
                <div class="metric-num">{total_weight}<span class="metric-unit">kg</span></div>
            </div>
            <div class="metric-card">
                <div class="metric-title">外箱总体积</div>
                <div class="metric-num">{total_vol}<span class="metric-unit">m³</span></div>
            </div>
            <div class="metric-card">
                <div class="metric-title">总体积重 (1:167)</div>
                <div class="metric-num">{total_vol_weight}<span class="metric-unit">kg</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # 2. 默认折叠的数据预览区域（保持页面整体紧凑简洁）
        with st.expander("📊 查看数据明细预览（点击展开 / 折叠）", expanded=False):
            tab1, tab2 = st.tabs(["汇总结果", "合并明细预览"])

            with tab1:
                st.dataframe(
                    summary_df,
                    use_container_width=True,
                    height=320
                )

            with tab2:
                st.dataframe(
                    merged_df.head(100),
                    use_container_width=True,
                    height=320
                )

        st.write("")

        # 3. 居中导出按钮
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        download_filename = f"发货单处理结果_{timestamp}.xlsx"

        col_left, col_btn, col_right = st.columns(3)
        with col_btn:
            st.download_button(
                label="⬇️ 导出处理后的 Excel 文件",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"处理发生异常：{str(e)}")