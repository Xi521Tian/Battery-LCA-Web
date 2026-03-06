import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import datetime
import re

st.set_page_config(page_title="电池LCA在线测算", page_icon="🔋", layout="wide")
st.title("🔋 动力电池全生命周期 (LCA) 在线核算与报告系统")
st.markdown("---")

# ==========================================
# 🌟 第一步：配置报告基础信息 (彻底补全 Part1 & Part2)
# ==========================================
st.sidebar.header("📝 填报导航")
st.sidebar.info("请依次完成基础信息配置与数据录入，系统将自动生成 6 大章节及附录的完整报告。")

st.header("第一步：配置报告基础信息")

# --- Part 1 概述 ---
with st.expander("📂 Part 1: 概述 (需手动填写)", expanded=True):
    company_intro = st.text_area("1.1 企业简介", placeholder="请输入企业基本情况...")
    product_intro = st.text_area("1.2 产品介绍", placeholder="请输入产品型号、规格...")
    production_process = st.text_area("1.3 产品生产工艺", placeholder="简述生产制造流程...")

    def_p1_4 = "基于LCA的评价方法，本产品碳足迹评价报告主要依据了以下标准：国际标准化组织（ISO）发布的《ISO 14067：2018温室气体—产品碳足迹—量化要求与指南》；英国标准协会（BSI）发布的《PAS 2050：2011商品和服务生命周期温室气体排放评价规范》；世界资源研究所（WRI）发布的《温室气体核算体系：产品寿命周期核算与报告标准》。\n依据产品所属行业标准对产品种类规则（Product Category Rules, PCR）的要求，根据体验账号的产品管理规定来定义产品的功能单位、边界、分配等计算原则。"
    evaluation_basis = st.text_area("1.4 评价依据和要求", value=def_p1_4)

# --- Part 2 目的和范围定义 ---
def_p2_1 = "本报告旨在通过揭示体验账号生产的汽车动力电池全生命周期的产品碳足迹，为体验账号持续开展节能减排工作提供数据参考；为体验账号向价值链相关方开展碳信息披露提供重要内容。"
def_p2_2 = "本次研究的对象为“汽车动力电池，100kwh”，为了方便输入/输出的量化，保证碳足迹分析结果的可比性，功能单位定义为“1个汽车动力电池”。"
def_p2_3 = "本次评价汽车动力电池的产品碳足迹，其生产周期为2026年01月01日至2026年12月31日。"
def_p2_4 = "本次产品碳足迹评价的系统边界为全生命周期-从资源开采到产品废弃，碳足迹核算包括原材料获取阶段、生产制造阶段、分销和储存阶段、产品使用阶段、废弃处置阶段的温室气体排放或清除。\n本次评估不涉及土地利用变化所导致的温室气体排放或清除，以及产品生物固碳。"
def_p2_5 = "本产品碳足迹评价报告核算二氧化碳（CO2）、甲烷（CH4）、氧化亚氮（N2O）、氢氟碳化物（HFCs）、全氟碳化物（PFCs）、六氟化硫（SF6）和三氟化氮（NF3）向大气中的排放或清除，采用IPCC第六次评估报告100-year的GWP值，将不同温室气体折算为二氧化碳当量（CO2e）。"
def_p2_6 = "本次产品碳足迹评价包括系统边界内所有对产品生命周期温室气体排放具有实质性贡献的排放源。本报告采取的数据取舍原则，以原材料投入占产品重量的比例、排放贡献占总排放的比例为依据，具体原则如下：\n（1）普通物料重量小于产品重量1%时，含有稀贵金属的物料小于产品重量0.1%时，可以忽略，但总共忽略的物料不超过产品重量的5%。\n（2）对于一些较难获取的数据，若其对于排放的影响很小，可对其进行忽略。\n（3）仅考虑生产过程的排放，生活过程不予考虑。\n（4）优先采用供应量最大的供应商数据。\n（5）运输距离无法获取时，采用平均运输距离。"

with st.expander("📂 Part 2: 目的和范围定义 (不填则自动使用默认标准文本)", expanded=False):
    purpose = st.text_area("2.1 研究目的", placeholder=def_p2_1)
    functional_unit = st.text_area("2.2 功能单位", placeholder=def_p2_2)
    time_scope = st.text_area("2.3 时间范围", placeholder=def_p2_3)
    system_boundary = st.text_area("2.4 系统边界", placeholder=def_p2_4)
    ghg_types = st.text_area("2.5 温室气体种类", placeholder=def_p2_5)
    cutoff_rules = st.text_area("2.6 取舍原则", placeholder=def_p2_6)

st.markdown("---")
st.header("第二步：录入生命周期测算数据")

# ==========================================
# 🌟 因子库与 UI 结构
# ==========================================
FACTOR_DB = {
    "天然气 (m3)": 2.0667, "厂务电力 (kWh)": 0.6205, "水 (m3)": 0.344, "废水 (t)": 0.118,
    "正极材料-磷酸铁锂 (kg)": 25.0, "负极材料-石墨 (kg)": 5.5, "电解液 (kg)": 19.6, "隔膜 (kg)": 3.24,
    "铝箔 (kg)": 2.39, "铜箔 (kg)": 12.4, "电池壳体-铝合金 (kg)": 28.38, "电池管理系统BMS (kg)": 28.38,
    "冷却系统-水冷板 (kg)": 11.5, "连接件-铜排 (kg)": 12.4, "电解质锂盐 (kg)": 0.01, "导电剂 (kg)": 3.9,
    "粘结剂 (kg)": 0.002, "电池结构胶 (kg)": 28.38, "绝缘材料 (kg)": 1.85,
    "木质托盘 (kg)": 1.28, "塑料薄膜 (kg)": 3.17, "纸质护角 (个)": 1.13,
    "光伏发电电力 (kWh)": 0.65, "电芯涂布烘干耗电 (kWh)": 0.6205, "辊压工序耗电 (kWh)": 0.6205,
    "电芯注液耗电 (kWh)": 0.6205, "化成分容耗电 (kWh)": 0.6205, "模组焊接耗电 (kWh)": 0.6205,
    "Pack装配耗电 (kWh)": 0.6205, "有机废气 (m3)": 0.056,
    "仓储温控耗电 (kWh)": 0.6205, "车辆行驶充电耗电 (kWh)": 0.6205, "电池热管理系统耗电 (kWh)": 0.6205,
    "回收清洗废水 (t)": 0.858, "动力电池回收拆解耗电 (kWh)": 0.6205,
    "废弃正极材料 (kg)": 1.82, "废弃负极材料 (kg)": 1.82, "废弃电解液 (kg)": 19.6, "废弃隔膜 (kg)": 0.39,
    "废弃铝箔 (kg)": 2.39, "废弃铜箔 (kg)": 12.4, "废弃电池结构胶 (kg)": 28.38, "废弃绝缘材料 (kg)": 1.82,
    "废弃BMS (kg)": 28.38, "废弃铝合金壳体 (kg)": 28.38, "废弃水冷板 (kg)": 0.167, "废弃铜排 (kg)": 0.14
}
default_trans_factor = 0.078

UI_STRUCTURE = {
    "3.5.1 原材料获取阶段": {
        "主辅材与包装": ["正极材料-磷酸铁锂 (kg)", "负极材料-石墨 (kg)", "电解液 (kg)", "隔膜 (kg)", "铝箔 (kg)",
                         "铜箔 (kg)", "电解质锂盐 (kg)", "导电剂 (kg)", "粘结剂 (kg)", "电池壳体-铝合金 (kg)",
                         "电池管理系统BMS (kg)", "冷却系统-水冷板 (kg)", "连接件-铜排 (kg)", "电池结构胶 (kg)",
                         "绝缘材料 (kg)", "木质托盘 (kg)", "塑料薄膜 (kg)", "纸质护角 (个)"],
        "运输": ["正极材料运输 (tkm)", "负极材料运输 (tkm)", "电解液运输 (tkm)", "隔膜运输 (tkm)", "铝箔运输 (tkm)",
                 "铜箔运输 (tkm)", "辅料运输 (tkm)"]
    },
    "3.5.2 生产制造阶段": {
        "物料与能耗": ["厂务电力 (kWh)", "天然气 (m3)", "水 (m3)", "废水 (t)", "有机废气 (m3)", "光伏发电电力 (kWh)",
                       "电芯涂布烘干耗电 (kWh)", "辊压工序耗电 (kWh)", "电芯注液耗电 (kWh)", "化成分容耗电 (kWh)",
                       "模组焊接耗电 (kWh)", "Pack装配耗电 (kWh)"],
        "运输": ["厂内周转运输 (tkm)"]
    },
    "3.5.3 分销和储存阶段": {
        "物料与能耗": ["仓储温控耗电 (kWh)"],
        "运输": ["电池仓储周转运输 (tkm)", "汽车动力电池成品运输 (tkm)"]
    },
    "3.5.4 产品使用阶段": {
        "物料与能耗": ["车辆行驶充电耗电 (kWh)", "电池热管理系统耗电 (kWh)"],
        "运输": ["售后维保运输 (tkm)"]
    },
    "3.5.5 废弃处置阶段": {
        "物料与能耗": ["动力电池回收拆解耗电 (kWh)", "回收清洗废水 (t)", "废弃正极材料 (kg)", "废弃负极材料 (kg)",
                       "废弃电解液 (kg)", "废弃隔膜 (kg)", "废弃铝箔 (kg)", "废弃铜箔 (kg)", "废弃电池结构胶 (kg)",
                       "废弃绝缘材料 (kg)", "废弃BMS (kg)", "废弃铝合金壳体 (kg)", "废弃水冷板 (kg)", "废弃铜排 (kg)"],
        "运输": ["废弃电池运输 (tkm)", "废弃包装材料运输 (tkm)"]
    }
}

user_records = {}
results = {stage: 0.0 for stage in UI_STRUCTURE.keys()}

for stage_name, categories in UI_STRUCTURE.items():
    user_records[stage_name] = {"Material": [], "Transport": []}
    with st.expander(f"⚙️ 展开录入：{stage_name}", expanded=False):
        for cat_name, items in categories.items():
            st.markdown(f"**📍 {cat_name}**")
            cols = st.columns(3)
            for i, item in enumerate(items):
                with cols[i % 3]:
                    val = st.number_input(item, min_value=0.0, step=1.0, key=f"{stage_name}_{item}")
                    match = re.search(r"(.+)\s*\((.+)\)", item)
                    name = match.group(1).strip() if match else item
                    unit = match.group(2).strip() if match else "-"
                    if "运输" in cat_name or "运输" in name:
                        factor = FACTOR_DB.get(item, default_trans_factor)
                        user_records[stage_name]["Transport"].append((name, val, unit, "运输"))
                    else:
                        factor = FACTOR_DB.get(item, 0.0)
                        user_records[stage_name]["Material"].append((name, val, unit, "主料/辅料/能耗"))
                    results[stage_name] += val * factor
            st.divider()


# Word 表格辅助函数
def add_word_table(doc, title, headers, data_rows):
    doc.add_paragraph(title)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    for i, h in enumerate(headers): table.rows[0].cells[i].text = h
    for r in data_rows:
        row = table.add_row().cells
        for i, val in enumerate(r): row[i].text = str(val)


# ==========================================
# 🌟 一键生成深度还原的 Word 报告 (全章节不打折)
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 提交全部数据，生成完整规范报告", type="primary", use_container_width=True):
    total_carbon = sum(results.values())
    st.success("✅ 报告生成完毕！本次包含了完整的 1-6 章节及附录。")

    doc = Document()

    # 封面
    doc.add_heading('产品碳足迹评价报告', 0).alignment = 1
    doc.add_paragraph('——汽车动力电池，100kwh\n').alignment = 1
    doc.add_paragraph(f'报告单位： 体验账号\n编制日期： {datetime.date.today().strftime("%Y年%m月%d日")}\n')

    # ---- 修复：Part 1 完全输出 ----
    doc.add_heading('1. 概述', level=1)
    doc.add_heading('1.1 企业简介', level=2)
    doc.add_paragraph(company_intro if company_intro.strip() else "（待补充）")
    doc.add_heading('1.2 产品介绍', level=2)
    doc.add_paragraph(product_intro if product_intro.strip() else "（待补充）")
    doc.add_heading('1.3 产品生产工艺', level=2)
    doc.add_paragraph(production_process if production_process.strip() else "（待补充）")
    doc.add_heading('1.4 评价依据和要求', level=2)
    doc.add_paragraph(evaluation_basis if evaluation_basis.strip() else def_p1_4)

    # ---- 修复：Part 2 完全输出 ----
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

    # ---- Part 3 数据清单输出 ----
    doc.add_heading('3. 生命周期数据清单分析', level=1)
    doc.add_heading('3.1 数据质量要求', level=2)
    doc.add_paragraph(
        '综合ISO 14067：2018、PAS 2050：2011对数据质量的要求，本次评价产品碳足迹的活动水平和排放因子满足相关性、完整性、一致性、连贯性、透明性、准确性及避免重复计算的要求。')
    doc.add_heading('3.2 碳足迹计算方法', level=2)
    doc.add_paragraph('采用公式：E = Σ (AD × EF × GWP) 进行计算。')
    doc.add_heading('3.3 分配', level=2)
    doc.add_paragraph(
        '在无法避免分配的情况下，系统的输入和输出应以反应它们之间潜在的物理关系的方式，在其不同的产品或功能之间进行划分。')
    doc.add_heading('3.4 假设', level=2)
    doc.add_paragraph('基于数据可得性与碳足迹核算需求的综合评估，已在合理假设下完成测算。')
    doc.add_heading('3.5 生命周期清单数据', level=2)

    table_idx = 2
    for stage_name, records in user_records.items():
        # 这里仅提取阶段名称纯文字，比如把 "3.5.1 原材料获取阶段" 变成 "原材料获取阶段" 方便做表格标题
        pure_stage_name = stage_name.split(' ')[1]
        doc.add_heading(stage_name, level=3)
        mat_data = records["Material"]
        if mat_data:
            doc.add_paragraph(f'{pure_stage_name}的清单数据，如表3-{table_idx}所示。')
            add_word_table(doc, f'表 3-{table_idx}：{pure_stage_name}数据清单', ['名称', '消耗数量', '单位', '类型'],
                           mat_data)
            table_idx += 1
        trans_data = records["Transport"]
        if trans_data:
            trans_formatted = [[row[0], "1", "道路运输", "-", row[1], "-"] for row in trans_data]
            doc.add_paragraph(f'{pure_stage_name}的运输清单数据，如表3-{table_idx}所示。')
            add_word_table(doc, f'表 3-{table_idx}：{pure_stage_name}运输数据清单',
                           ['运输物', '路段', '运输方式', '货物重量', '运输里程', '能源消耗'], trans_formatted)
            table_idx += 1

    # ---- Part 4 结果 ----
    doc.add_heading('4. 产品碳足迹评价结果', level=1)
    doc.add_paragraph(
        f'通过计算，功能单位（1个汽车动力电池）的全生命周期碳足迹为 {total_carbon:,.6f} kgCO2e，单位产品排放量为 {total_carbon:,.6f} kgCO2e/个，碳足迹的整体情况，如表4-1所示。')
    res_data = []
    for stage, em in results.items():
        pct = f"{(em / total_carbon * 100):.4f}" if total_carbon > 0 else "0.0000"
        res_data.append([stage.split(" ")[1], f"{em:,.6f}", pct])
    res_data.append(['总计', f"{total_carbon:,.6f}", '100'])
    add_word_table(doc, '表4-1：产品碳足迹评价结果', ['生命周期阶段', '排放量（kgCO2e）', '占比（%）'], res_data)

    # ---- 修复：Part 5 满血输出所有的表格 ----
    doc.add_heading('5. 不确定性分析', level=1)
    doc.add_heading('5.1 不确定性分析方法', level=2)
    doc.add_paragraph(
        '产品碳足迹评价的不确定性分析，采用定性分析法，包括活动水平数据质量评级、排放因子的质量评级。活动水平数据质量评级，如表5-1所示。')

    add_word_table(doc, '表5-1：活动水平数据质量评级', ['质量等级', '描述', '分值'],
                   [['好', '量测值：实际量测数值...', '5'], ['较好', '计算值：以某合理方法进行计算的数值...', '3'],
                    ['一般', '理论值/经验值：根据理论推导...', '2'], ['差', '参考文献：由其它文献取得...', '1']])

    doc.add_paragraph(
        '排放因子质量评级，从时间相关性、地域相关性、技术相关性、数据准确度、方法学等方面评定，具体标准，如表5-2、5-3、5-4、5-5、5-6所示。')
    add_word_table(doc, '表5-2：排放因子的质量评级-时间相关性', ['时间相关性', '分值'],
                   [['<5年', '5'], ['5–10年', '3'], ['10–15年', '2'], ['>15年（及未知年份）', '1']])
    add_word_table(doc, '表5-3：排放因子的质量评级-地域相关性', ['地域相关性', '分值'],
                   [['完全符合所评估产品生产地点', '5'], ['数据为国家层面的数据', '3'], ['数据为全球平均数据', '1']])
    add_word_table(doc, '表5-4：排放因子的质量评级-技术相关性', ['技术相关性', '分值'],
                   [['完全符合所评估产品生产技术', '5'], ['行业平均数据', '3'], ['替代数据', '1']])
    add_word_table(doc, '表5-5：排放因子的质量评级-数据准确度', ['数据准确度', '分值'],
                   [['变异性低', '5'], ['变异性未量化，考虑为较低', '3'], ['变异性未量化，考虑为较高', '2'],
                    ['变异性高', '1']])
    add_word_table(doc, '表5-6：排放因子的质量评级-方法学', ['方法学的适合性及一致性', '分值'],
                   [['PAS 2050/补充要求所规定的排放因子', '5'], ['政府/国际政府组织/行业发布的排放因子', '3'],
                    ['公司/其他机构发布的排放因子', '1']])

    doc.add_heading('5.2 不确定性分析结果', level=2)
    doc.add_paragraph('汽车动力电池产品碳足迹评价的活动数据和排放因子数据质量分析结果，如表5-7、5-8所示。')
    add_word_table(doc, '表5-7：活动数据质量分析结果', ['活动数据类别', '活动数据描述', '质量级别', '得分'],
                   [['原材料类', '', '', ''], ['能源资源类', '', '', ''], ['运输类', '', '', ''],
                    ['产品使用类', '', '', ''], ['废弃处置类', '', '', '']])
    add_word_table(doc, '表5-8：排放因子数据质量分析结果', ['排放因子类别', '排放因子描述', '平均得分'],
                   [['原材料类', '', ''], ['能源资源类', '', ''], ['运输类', '', ''], ['产品使用类', '', ''],
                    ['废弃处置类', '', '']])

    # ---- Part 6 结论 ----
    doc.add_heading('6. 结论', level=1)
    pcts = [(em / total_carbon * 100) if total_carbon > 0 else 0 for em in results.values()]
    doc.add_paragraph(
        f'体验账号汽车动力电池的产品碳足迹评价，涵盖的时间范围是2026年01月01日至2026年12月31日。功能单位（1个汽车动力电池）的全生命周期碳足迹为 {total_carbon:,.6f} kgCO2e，其中原材料获取阶段、生产制造阶段、分销和储存阶段、产品使用阶段、废弃处置阶段的排放占比分别为：{pcts[0]:.2f}%、{pcts[1]:.2f}%、{pcts[2]:.2f}%、{pcts[3]:.2f}%、{pcts[4]:.2f}%。')

    # ---- 修复：附录完整输出 ----
    doc.add_heading('附录：排放因子选择表', level=1)
    appendix_data = [
        ['天然气', '天然气', '2.06672', 'kgCO2e/m3', 'Department for Environment'],
        ['厂务电力', '电力', '0.6205', 'kgCO2e/kWh', 'Ministry of Ecology and Environment'],
        ['自来水消耗', '自来水', '0.344', 'kgCO2e/m3', 'Department for Environment'],
        ['正极材料(磷酸铁锂)', '铸铁材料', '1.82', 'kgCO2e/kg', 'China Products Carbon Footprint Factors Database'],
        ['负极材料(石墨)', '石墨', '5.5', 'kgCO2e/kg', 'China Automotive Data Co. Ltd'],
        ['电解液', '电解液：六氟磷酸锂', '19.6', 'kgCO2e/kg', 'China Automotive Data Co. Ltd'],
        ['隔膜', '塑料薄膜包装袋', '3.24', 'kgCO2e/kg', 'China Products Carbon Footprint Factors Database'],
        ['铝箔', '铝箔', '2.39', 'kgCO2e/kg', 'Korea Environmental Industry Technology Research Institute'],
        ['铜箔', '铜箔', '12.4', 'kgCO2e/kg', 'Taiwan Environmental Protection Agency'],
        ['电池壳体(铝合金)', '电池-镍氢电池', '28380', 'kgCO2e/t', 'Department for Environment'],
        ['木质托盘', '木质托盘48x40', '1.28', 'kgCO2e/kg', 'Scientific Data'],
        ['公路物流运输', '重型柴油货车运输（载重30t）', '0.078', 'kgCO2e/tkm', '中华人民共和国住房和城乡建设部']
    ]
    add_word_table(doc, '附表：系统引用的主要排放因子',
                   ['排放源名称', '排放因子名称', '排放因子数值', '排放因子单位', '排放因子来源'], appendix_data)
    doc.add_paragraph(
        '参考文献：\n[1] 《商品和服务在生命周期内的温室气体排放评价规范》（PAS 2050:2011）\n[2] 《温室气体—产品碳足迹—量化要求与指南》（ISO 14067:2018）')

    bio = io.BytesIO()
    doc.save(bio)
    st.download_button(
        label="📥 点击下载完整 6 章节及附录 Word 测算报告",
        data=bio.getvalue(),
        file_name="产品碳足迹评价报告_极致完整版.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )