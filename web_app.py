import streamlit as st
import pandas as pd
from docx import Document
import io
import datetime  # 用于自动获取当天日期

st.set_page_config(page_title="电池LCA在线测算", page_icon="🔋", layout="wide")
st.title("🔋 动力电池全生命周期 (LCA) 在线核算与报告系统")
st.markdown("---")

# ==========================================
# 🌟 新增模块：报告基础信息配置 (Part 1 & Part 2)
# ==========================================
st.sidebar.header("📝 填报导航")
st.sidebar.info("请先完成报告基础信息填写，再进行各阶段碳排放数据录入。")

st.header("第一步：配置报告基础信息")

# --- Part 1 概述 (前端输入) ---
with st.expander("📂 Part 1: 概述 (需手动填写)", expanded=True):
    st.markdown("请在此填写企业及产品的基础信息，这些内容将直接输出至正式报告的第 1 章。")
    company_intro = st.text_area("1.1 企业简介", placeholder="请输入企业基本情况...")
    product_intro = st.text_area("1.2 产品介绍", placeholder="请输入该款动力电池的型号、规格、主要参数...")
    production_process = st.text_area("1.3 产品生产工艺", placeholder="简述产品的生产制造流程...")
    evaluation_basis = st.text_area("1.4 评价依据和要求",
                                    value="基于LCA的评价方法，本产品碳足迹评价报告主要依据了以下标准：国际标准化组织（ISO）发布的《ISO 14067：2018温室气体—产品碳足迹—量化要求与指南》；英国标准协会（BSI）发布的《PAS 2050：2011商品和服务生命周期温室气体排放评价规范》；世界资源研究所（WRI）发布的《温室气体核算体系：产品寿命周期核算与报告标准》。\n依据产品所属行业标准对产品种类规则（Product Category Rules, PCR）的要求，根据体验账号的产品管理规定来定义产品的功能单位、边界、分配等计算原则。")

# --- Part 2 目的和范围定义 (默认文本兜底机制) ---
# 以下是提取自你提供的标准模板的默认文本
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
# (原有代码保留区：因子库与数据计算逻辑)
# ==========================================
FACTOR_DB = {
    "天然气 (m3)": 2.162, "用电 (kWh)": 0.5703,
    "氢氧化锂/前驱体 (kg)": 16.50, "石墨 (kg)": 4.50, "电解液原料 (kg)": 8.20,
    "隔膜原料 (kg)": 2.80, "铝合金/钢材 (kg)": 11.50, "铜/铝箔等 (kg)": 9.50,
    "BMS及电子元器件 (kg)": 25.00, "包装材料 (kg)": 1.50,
    "公路运输周转量 (tkm)": 0.107,
    "NMP挥发逃逸量 (kg)": 2.50,
    "回收有价金属 (kg)": -8.50
}

UI_STRUCTURE = {
    "材料获取阶段": {
        "正极材料生产": ["天然气 (m3)", "用电 (kWh)", "氢氧化锂/前驱体 (kg)", "公路运输周转量 (tkm)"],
        "负极材料生产": ["天然气 (m3)", "用电 (kWh)", "石墨 (kg)", "公路运输周转量 (tkm)"]
    },
    "生产制造阶段": {
        "极片制造(搅拌/涂布)": ["天然气 (m3)", "用电 (kWh)", "NMP挥发逃逸量 (kg)"]
    }
    # 为了演示精简了部分代码，你可以把之前的完整字典贴回来
}

results = {stage: 0.0 for stage in UI_STRUCTURE.keys()}

for stage_name, processes in UI_STRUCTURE.items():
    with st.expander(f"⚙️ 数据录入：{stage_name}", expanded=False):
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

# ==========================================
# 🌟 重构：一键导出极度还原模板的 Word 报告
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 提交全部数据，生成符合规范的 LCA 碳足迹报告", type="primary", use_container_width=True):
    st.success("✅ 数据核算完成！文件已准备就绪。")
    total_carbon = sum(results.values())

    # 开始构建 Word 文档
    doc = Document()

    # ---- 封面及抬头 ----
    doc.add_heading('产品碳足迹评价报告', 0)
    doc.add_paragraph('——汽车动力电池，100kwh\n')
    doc.add_paragraph(f'报告单位： 体验账号')
    today_str = datetime.date.today().strftime("%Y年%m月%d日")
    doc.add_paragraph(f'编制日期： {today_str}\n')

    # ---- Part 1: 概述 ----
    doc.add_heading('1. 概述', level=1)

    doc.add_heading('1.1 企业简介', level=2)
    doc.add_paragraph(company_intro if company_intro.strip() else "（待补充）")

    doc.add_heading('1.2 产品介绍', level=2)
    doc.add_paragraph(product_intro if product_intro.strip() else "（待补充）")

    doc.add_heading('1.3 产品生产工艺', level=2)
    doc.add_paragraph(production_process if production_process.strip() else "（待补充）")

    doc.add_heading('1.4 评价依据和要求', level=2)
    doc.add_paragraph(evaluation_basis if evaluation_basis.strip() else "（待补充）")

    # ---- Part 2: 目的和范围定义 (使用逻辑判断兜底) ----
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

    # ---- Part 3 (演示结果，后续再按你的模板迭代表格) ----
    doc.add_heading('附：本次测算总计', level=1)
    doc.add_paragraph(f'生命周期总碳足迹：{total_carbon:,.2f} kg CO2e')

    # 将文件写入内存流并生成下载按钮
    bio = io.BytesIO()
    doc.save(bio)

    st.download_button(
        label="📥 点击下载标准 Word 测算报告",
        data=bio.getvalue(),
        file_name=f"汽车动力电池LCA评价报告_{today_str}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )