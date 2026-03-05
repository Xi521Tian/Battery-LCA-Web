import streamlit as st
import pandas as pd
from docx import Document  # 新增：用于生成 Word 文档
import io  # 新增：用于在内存中处理文件下载

# 1. 页面基本设置
st.set_page_config(page_title="电池LCA在线测算", page_icon="🔋", layout="wide")
st.title("🔋 动力电池全生命周期 (LCA) 在线核算系统")
st.markdown("---")

# 2. 核心大数据库 (因子库)
FACTOR_DB = {
    "天然气 (m3)": 2.162, "用电 (kWh)": 0.5703,
    "氢氧化锂/前驱体 (kg)": 16.50, "石墨 (kg)": 4.50, "电解液原料 (kg)": 8.20,
    "隔膜原料 (kg)": 2.80, "铝合金/钢材 (kg)": 11.50, "铜/铝箔等 (kg)": 9.50,
    "BMS及电子元器件 (kg)": 25.00, "包装材料 (kg)": 1.50,
    "公路运输周转量 (tkm)": 0.107,
    "NMP挥发逃逸量 (kg)": 2.50,
    "回收有价金属 (kg)": -8.50
}

# 3. 界面结构字典
UI_STRUCTURE = {
    "材料获取阶段": {
        "正极材料生产": ["天然气 (m3)", "用电 (kWh)", "氢氧化锂/前驱体 (kg)", "公路运输周转量 (tkm)"],
        "负极材料生产": ["天然气 (m3)", "用电 (kWh)", "石墨 (kg)", "公路运输周转量 (tkm)"],
        "电解液与隔膜": ["电解液原料 (kg)", "隔膜原料 (kg)", "用电 (kWh)", "公路运输周转量 (tkm)"],
        "外壳与BMS": ["铝合金/钢材 (kg)", "铜/铝箔等 (kg)", "BMS及电子元器件 (kg)", "公路运输周转量 (tkm)"]
    },
    "生产制造阶段": {
        "极片制造(搅拌/涂布)": ["天然气 (m3)", "用电 (kWh)", "NMP挥发逃逸量 (kg)"],
        "电芯装配与化成": ["用电 (kWh)", "包装材料 (kg)"]
    },
    "运输与使用阶段": {
        "成品物流分销": ["公路运输周转量 (tkm)"],
        "使用阶段损耗": ["用电 (kWh)"]
    },
    "回收处置阶段": {
        "电池拆解与湿法冶金": ["天然气 (m3)", "用电 (kWh)"],
        "再生材料产出 (碳抵扣)": ["回收有价金属 (kg)"]
    }
}

results = {stage: 0.0 for stage in UI_STRUCTURE.keys()}

# 4. 动态生成输入表单
st.sidebar.header("📝 填报说明")
st.sidebar.info("本系统按照 ISO 14067 标准构建，请依次展开各阶段填报实测数据。未发生项填 0 即可。")

for stage_name, processes in UI_STRUCTURE.items():
    with st.expander(f"📂 展开填报：{stage_name}", expanded=False):
        for process_name, inputs in processes.items():
            st.markdown(f"**📍 {process_name}**")
            cols = st.columns(3)
            for i, input_item in enumerate(inputs):
                with cols[i % 3]:
                    unique_key = f"{stage_name}_{process_name}_{input_item}"
                    user_val = st.number_input(input_item, min_value=0.0, step=10.0, key=unique_key)
                    factor = FACTOR_DB.get(input_item, 0.0)
                    results[stage_name] += user_val * factor
            st.divider()

# 5. 测算与报告导出逻辑
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 提交全部数据，生成 LCA 碳足迹报告", type="primary", use_container_width=True):
    st.success("✅ 数据核算完成！已生成多维数据报告。")

    total_carbon = sum(results.values())
    st.metric(label="🌟 动力电池全生命周期总碳足迹 (kgCO2e)", value=f"{total_carbon:,.2f}")

    st.markdown("### 📊 碳足迹分解视图")
    col_chart1, col_chart2 = st.columns(2)
    df_results = pd.DataFrame(list(results.items()), columns=["生命周期阶段", "碳排放量 (kgCO2e)"]).set_index(
        "生命周期阶段")

    with col_chart1:
        st.write("各阶段排放明细表")
        st.dataframe(df_results, use_container_width=True)
    with col_chart2:
        st.write("各阶段排放占比图")
        st.bar_chart(df_results)

    # ==========================================
    # 🌟 新增：后台排版并生成 Word 文档
    # ==========================================
    st.markdown("---")
    st.markdown("### 📄 导出正式报告文件")

    # 初始化 Word 文档
    doc = Document()
    doc.add_heading('动力电池全生命周期 (LCA) 碳足迹核算报告', 0)
    doc.add_paragraph('本报告由在线LCA系统根据您的填报数据自动生成。')

    doc.add_heading('一、 测算总计', level=1)
    doc.add_paragraph(f'生命周期总碳足迹：{total_carbon:,.2f} kg CO2e')

    doc.add_heading('二、 各阶段排放明细', level=1)
    # 创建 2 列的表格
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'  # 添加实线边框
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '生命周期阶段'
    hdr_cells[1].text = '碳排放量 (kg CO2e)'

    # 逐行填入数据
    for stage, emission in results.items():
        row_cells = table.add_row().cells
        row_cells[0].text = stage
        row_cells[1].text = f"{emission:,.2f}"

    # 将文件写入内存流，以便供用户下载 (不在服务器留存实体文件，非常安全)
    bio = io.BytesIO()
    doc.save(bio)

    # 在网页上生成一个华丽的下载按钮
    st.download_button(
        label="📥 点击下载 Word 格式测算报告",
        data=bio.getvalue(),
        file_name="电池LCA碳足迹测算报告.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )