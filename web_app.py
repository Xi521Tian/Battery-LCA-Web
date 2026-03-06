import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
import io
import datetime

st.set_page_config(page_title="电池LCA在线测算", page_icon="🔋", layout="wide")
st.title("🔋 动力电池全生命周期 (LCA) 在线核算与报告系统")
st.markdown("---")

# ==========================================
# 🌟 第一步：报告基础信息配置 (Part 1 & Part 2)
# ==========================================
st.sidebar.header("📝 填报导航")
st.sidebar.info("第一步：完成报告基础信息填写。\n第二步：展开各阶段进行数据录入。\n第三步：一键生成标准报告。")

st.header("第一步：配置报告基础信息")

with st.expander("📂 Part 1: 概述 (需手动填写)", expanded=True):
    st.markdown("请在此填写企业及产品的基础信息，这些内容将直接输出至正式报告的第 1 章。")
    company_intro = st.text_area("1.1 企业简介", placeholder="请输入企业基本情况...")
    product_intro = st.text_area("1.2 产品介绍", placeholder="请输入该款动力电池的型号、规格、主要参数...")
    production_process = st.text_area("1.3 产品生产工艺", placeholder="简述产品的生产制造流程...")
    evaluation_basis = st.text_area("1.4 评价依据和要求",
                                    value="基于LCA的评价方法，本产品碳足迹评价报告主要依据了以下标准：国际标准化组织（ISO）发布的《ISO 14067：2018温室气体—产品碳足迹—量化要求与指南》；英国标准协会（BSI）发布的《PAS 2050：2011商品和服务生命周期温室气体排放评价规范》；世界资源研究所（WRI）发布的《温室气体核算体系：产品寿命周期核算与报告标准》。\n依据产品所属行业标准对产品种类规则（Product Category Rules, PCR）的要求，根据体验账号的产品管理规定来定义产品的功能单位、边界、分配等计算原则。")

def_p2_1 = "本报告旨在通过揭示体验账号生产的汽车动力电池全生命周期的产品碳足迹，为体验账号持续开展节能减排工作提供数据参考；为体验账号向价值链相关方开展碳信息披露提供重要内容。"
def_p2_2 = "本次研究的对象为“汽车动力电池，100kwh”，为了方便输入/输出的量化，保证碳足迹分析结果的可比性，功能单位定义为“1个汽车动力电池”。"
def_p2_3 = "本次评价汽车动力电池的产品碳足迹，其生产周期为2026年01月01日至2026年12月31日。"
def_p2_4 = "本次产品碳足迹评价的系统边界为全生命周期-从资源开采到产品废弃，碳足迹核算包括原材料获取阶段、生产制造阶段、分销和储存阶段、产品使用阶段、废弃处置阶段的温室气体排放或清除。\n本次评估不涉及土地利用变化所导致的温室气体排放或清除，以及产品生物固碳。"
def_p2_5 = "本产品碳足迹评价报告核算二氧化碳（CO2）、甲烷（CH4）、氧化亚氮（N2O）、氢氟碳化物（HFCs）、全氟碳化物（PFCs）、六氟化硫（SF6）和三氟化氮（NF3）向大气中的排放或清除，采用IPCC第六次评估报告100-year的GWP值，将不同温室气体折算为二氧化碳当量（CO2e）。"
def_p2_6 = "本次产品碳足迹评价包括系统边界内所有对产品生命周期温室气体排放具有实质性贡献的排放源。本报告采取的数据取舍原则，以原材料投入占产品重量的比例、排放贡献占总排放的比例为依据，具体原则如下：\n（1）普通物料重量小于产品重量1%时，含有稀贵金属(如金、银、铂、钯等)或高纯物质(如纯度高于99.99%)的物料小于产品重量0.1%时，可以忽略，但总共忽略的物料不超过产品重量的5%。\n（2）对于一些较难获取的数据，若其对于排放的影响很小（单个影响低于1%），可对其进行忽略，忽略的数据或阶段的排放量之和不应超过10%。\n（3）仅考虑生产过程的排放，对于厂区生活过程如食堂等的排放不予考虑；当生活与生产的排放无法区分时，将合并考虑。\n（4）在计算原材料获取阶段的排放量时，若原材料来源于多个供应商，采用供应量最大的供应商提供的数据；若最大的供应商数据无法获取时，将采用供应量第二大的供应商提供的数据，依此原则逐步获取数据；若所有供应商数据均无法获取，则采用行业默认的排放数据。\n（5）在计算原材料运输的排放量时，若实际原料的运输距离数据不可获得时，采用供应量最大的供应商的平均运输距离。"

with st.expander("📂 Part 2: 目的和范围定义 (选填，不填则默认使用标准模板)", expanded=False):
    st.markdown("如需自定义，请在下方修改；若留空，系统将自动填入标准的合规话术。")
    purpose = st.text_area("2.1 研究目的", placeholder=def_p2_1)
    functional_unit = st.text_area("2.2 功能单位", placeholder=def_p2_2)
    time_scope = st.text_area("2.3 时间范围", placeholder=def_p2_3)
    system_boundary = st.text_area("2.4 系统边界", placeholder=def_p2_4)
    ghg_types = st.text_area("2.5 温室气体种类", placeholder=def_p2_5)
    cutoff_rules = st.text_area("2.6 取舍原则", placeholder=def_p2_6)

st.markdown("---")
st.header("第二步：录入生命周期测算数据")

# ==========================================
# 🌟 第二步：扩充版因子库与高度还原的 UI 结构
# ==========================================
# 因子库严格参照报告附录的排放系数进行匹配
FACTOR_DB = {
    # 主辅材 (kg)
    "正极材料-磷酸铁锂 (kg)": 25.0, "负极材料-石墨 (kg)": 5.5, "电解液 (kg)": 19.6,
    "隔膜 (kg)": 3.24, "铝箔 (kg)": 2.39, "铜箔 (kg)": 12.4,
    "电解质锂盐 (kg)": 0.01, "导电剂 (kg)": 3.9, "粘结剂 (kg)": 0.00223,
    "电池壳体-铝合金 (kg)": 28.38, "电池管理系统BMS (kg)": 28.38, "冷却系统-水冷板 (kg)": 11.5,
    "连接件-铜排 (kg)": 12.4, "电池结构胶 (kg)": 28.38, "绝缘材料 (kg)": 1.85,
    # 包装材
    "木质托盘 (kg)": 1.28, "塑料薄膜 (kg)": 3.17, "纸质护角 (个)": 1.131,
    # 运输类
    "重型柴油卡车运输 (tkm)": 0.078, "铁路运输 (tkm)": 0.0278,
    # 能源与排放
    "天然气 (m3)": 2.066, "外购电力 (kWh)": 0.6205, "工业废水处理 (t)": 0.118,
    "自来水消耗 (m3)": 0.344, "NMP挥发逃逸量 (kg)": 2.50
}

# 数据驱动的 UI 结构：字典不仅决定算什么，也决定网页长什么样
UI_STRUCTURE = {
    "材料获取阶段": {
        "核心电芯主材": [
            "正极材料-磷酸铁锂 (kg)", "负极材料-石墨 (kg)", "电解液 (kg)",
            "隔膜 (kg)", "铝箔 (kg)", "铜箔 (kg)",
            "电解质锂盐 (kg)", "导电剂 (kg)", "粘结剂 (kg)"
        ],
        "Pack结构件与辅材": [
            "电池壳体-铝合金 (kg)", "电池管理系统BMS (kg)", "冷却系统-水冷板 (kg)",
            "连接件-铜排 (kg)", "电池结构胶 (kg)", "绝缘材料 (kg)"
        ],
        "产品包装材料": [
            "木质托盘 (kg)", "塑料薄膜 (kg)", "纸质护角 (个)"
        ],
        "原材料物流运输": [
            "重型柴油卡车运输 (tkm)", "铁路运输 (tkm)"
        ]
    },
    "生产制造阶段": {
        "动力设备与厂务": ["外购电力 (kWh)", "天然气 (m3)", "自来水消耗 (m3)", "工业废水处理 (t)"],
        "极片与电芯制造": ["NMP挥发逃逸量 (kg)"]
    },
    "分销和储存阶段": {
        "成品物流分销": ["重型柴油卡车运输 (tkm)"],
        "仓储环节耗电": ["外购电力 (kWh)"]
    },
    "产品使用阶段": {
        "运行期能量补给": ["外购电力 (kWh)"]
    },
    "废弃处置阶段": {
        "报废物流回收": ["重型柴油卡车运输 (tkm)"],
        "拆解与无害化处理": ["外购电力 (kWh)", "天然气 (m3)"]
    }
}

results = {stage: 0.0 for stage in UI_STRUCTURE.keys()}

# 动态生成输入框网格
for stage_name, processes in UI_STRUCTURE.items():
    with st.expander(f"⚙️ 数据录入：{stage_name}", expanded=False):
        for process_name, inputs in processes.items():
            st.markdown(f"**📍 {process_name}**")
            cols = st.columns(3)  # 排列成整齐的三列
            for i, input_item in enumerate(inputs):
                with cols[i % 3]:
                    unique_key = f"{stage_name}_{process_name}_{input_item}"
                    # 界面上显示物料名称，用户输入数量
                    user_val = st.number_input(input_item, min_value=0.0, step=1.0, key=unique_key)
                    # 匹配因子并累加碳排
                    factor = FACTOR_DB.get(input_item, 0.0)
                    results[stage_name] += user_val * factor
            st.divider()

# ==========================================
# 🌟 第三步：一键生成排版精美的 Word 报告
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 提交全部数据，生成符合规范的 LCA 碳足迹报告", type="primary", use_container_width=True):
    total_carbon = sum(results.values())

    st.success("✅ 数据核算完成！文件已准备就绪，请点击下方按钮下载。")
    st.metric(label="🌟 功能单位生命周期总碳足迹 (kgCO2e)", value=f"{total_carbon:,.2f}")

    # --- Word 文档构建 ---
    doc = Document()

    # 封面标题居中加粗
    title = doc.add_heading('产品碳足迹评价报告', 0)
    title.alignment = 1  # 居中
    doc.add_paragraph('——汽车动力电池，100kwh\n').alignment = 1

    today_str = datetime.date.today().strftime("%Y年%m月%d日")
    doc.add_paragraph(f'报告单位： 体验账号\n编制日期： {today_str}\n\n')

    # Part 1: 概述
    doc.add_heading('1. 概述', level=1)
    doc.add_heading('1.1 企业简介', level=2)
    doc.add_paragraph(company_intro if company_intro.strip() else "（待补充企业简要概况...）")
    doc.add_heading('1.2 产品介绍', level=2)
    doc.add_paragraph(product_intro if product_intro.strip() else "（待补充产品规格与参数...）")
    doc.add_heading('1.3 产品生产工艺', level=2)
    doc.add_paragraph(production_process if production_process.strip() else "（待补充工艺流程...）")
    doc.add_heading('1.4 评价依据和要求', level=2)
    doc.add_paragraph(evaluation_basis if evaluation_basis.strip() else "（待补充）")

    # Part 2: 目的和范围定义
    doc.add_heading('2. 目的和范围定义', level=1)
    doc.add_heading('2.1 研究目的', level=2)
    doc.add_paragraph(purpose if purpose.strip() else def_p2_1)
    doc.add_heading('2.2 功能单位', level=2)
    doc.add_paragraph(functional_unit if functional_unit.strip() else def_p2_2)
    doc.add_heading('2.3 时间范围', level=2)
    doc.add_paragraph(time_scope if time_scope.strip() else def_p2_3)
    doc.add_heading('2.4 系统边界', level=2)
    doc.add_paragraph(system_boundary if system_boundary.strip() else def_p2_4)
    doc.add_heading('2.5 温室气体种类', level=2)
    doc.add_paragraph(ghg_types if ghg_types.strip() else def_p2_5)
    doc.add_heading('2.6 取舍原则', level=2)
    doc.add_paragraph(cutoff_rules if cutoff_rules.strip() else def_p2_6)

    # Part 4: 评价结果表格输出
    doc.add_heading('4. 产品碳足迹评价结果', level=1)
    doc.add_paragraph(
        f'通过计算，功能单位（1个汽车动力电池）的全生命周期碳足迹为 {total_carbon:,.2f} kgCO2e，碳足迹的整体情况，如下表所示。')

    # 创建结果表格
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '生命周期阶段'
    hdr_cells[1].text = '排放量（kgCO2e）'
    hdr_cells[2].text = '占比（%）'

    for stage, emission in results.items():
        row_cells = table.add_row().cells
        row_cells[0].text = stage
        row_cells[1].text = f"{emission:,.2f}"
        percentage = (emission / total_carbon * 100) if total_carbon > 0 else 0
        row_cells[2].text = f"{percentage:.2f}%"

    # 添加合计行
    total_row = table.add_row().cells
    total_row[0].text = '总计'
    total_row[1].text = f"{total_carbon:,.2f}"
    total_row[2].text = '100%'

    # 写入内存供下载
    bio = io.BytesIO()
    doc.save(bio)

    st.download_button(
        label="📥 点击下载标准 Word 测算报告",
        data=bio.getvalue(),
        file_name=f"汽车动力电池LCA评价报告_{today_str}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )