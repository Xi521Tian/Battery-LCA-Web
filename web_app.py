import streamlit as st
import pandas as pd
from docx import Document
import io
import datetime
import re

st.set_page_config(page_title="电池LCA在线测算", page_icon="🔋", layout="wide")
st.title("🔋 动力电池全生命周期 (LCA) 在线核算与报告系统")
st.markdown("---")

# ==========================================
# 🌟 报告基础信息配置 (Part 1 & Part 2)
# ==========================================
st.sidebar.header("📝 填报导航")
st.sidebar.info("按照报告规范，请依次完成基础信息与清单数据录入。")

st.header("第一步：配置报告基础信息")

with st.expander("📂 Part 1 & 2: 概述与范围定义 (展开填写)", expanded=False):
    company_intro = st.text_area("1.1 企业简介", placeholder="请输入企业基本情况...")
    product_intro = st.text_area("1.2 产品介绍", placeholder="请输入产品型号、规格...")
    production_process = st.text_area("1.3 产品生产工艺", placeholder="简述生产制造流程...")

    # Part 2 默认兜底文案
    def_p2_1 = "本报告旨在通过揭示体验账号生产的汽车动力电池全生命周期的产品碳足迹，为体验账号持续开展节能减排工作提供数据参考；为体验账号向价值链相关方开展碳信息披露提供重要内容。"
    def_p2_2 = "本次研究的对象为“汽车动力电池，100kwh”，为了方便输入/输出的量化，保证碳足迹分析结果的可比性，功能单位定义为“1个汽车动力电池”。"
    def_p2_3 = "本次评价汽车动力电池的产品碳足迹，其生产周期为2026年01月01日至2026年12月31日。"
    purpose = st.text_area("2.1 研究目的", placeholder=def_p2_1)
    functional_unit = st.text_area("2.2 功能单位", placeholder=def_p2_2)
    time_scope = st.text_area("2.3 时间范围", placeholder=def_p2_3)

st.markdown("---")
st.header("第二步：录入生命周期测算数据")

# ==========================================
# 🌟 核心：全覆盖因子库与 UI 结构 (对照模板补全)
# ==========================================
# 提取自你附录的基准系数 (为了演示，电力统一取0.6205，运输取0.078)
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
# 设置默认的物流因子
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

# 结构化记录用户输入，用于生成 Word 第三部分的表格
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

                    # 提取名称和单位
                    match = re.search(r"(.+)\s*\((.+)\)", item)
                    name = match.group(1).strip() if match else item
                    unit = match.group(2).strip() if match else "-"

                    # 归类并累加计算
                    if "运输" in cat_name or "运输" in name:
                        factor = FACTOR_DB.get(item, default_trans_factor)
                        user_records[stage_name]["Transport"].append((name, val, unit, "运输"))
                    else:
                        factor = FACTOR_DB.get(item, 0.0)
                        user_records[stage_name]["Material"].append((name, val, unit, "主料/辅料/能耗"))

                    results[stage_name] += val * factor
            st.divider()


# 辅助函数：快速生成 Word 表格
def add_word_table(doc, title, headers, data_rows):
    doc.add_paragraph(title)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    for i, h in enumerate(headers): table.rows[0].cells[i].text = h
    for r in data_rows:
        row = table.add_row().cells
        for i, val in enumerate(r): row[i].text = str(val)


# ==========================================
# 🌟 一键生成深度还原的 Word 报告
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 提交全部数据，生成符合规范的 LCA 碳足迹报告", type="primary", use_container_width=True):
    total_carbon = sum(results.values())
    st.success("✅ 数据核算完成！请点击下方按钮下载完整版报告。")

    doc = Document()

    # 封面
    doc.add_heading('产品碳足迹评价报告', 0).alignment = 1
    doc.add_paragraph('——汽车动力电池，100kwh\n').alignment = 1
    doc.add_paragraph(f'报告单位： 体验账号\n编制日期： {datetime.date.today().strftime("%Y年%m月%d日")}\n')

    # Part 1 & 2 略写 (与上一版一致，写入输入的内容或兜底文本)
    doc.add_heading('1. 概述', level=1)
    doc.add_paragraph(company_intro if company_intro.strip() else "（待补充）")
    doc.add_heading('2. 目的和范围定义', level=1)
    doc.add_paragraph(purpose if purpose.strip() else def_p2_1)

    # ---- 核心升级：Part 3 生命周期数据清单分析 ----
    doc.add_heading('3. 生命周期数据清单分析', level=1)
    doc.add_paragraph(
        '3.1 数据质量要求\n（略，同标准模板）\n3.2 碳足迹计算方法\n（略，同标准模板）\n3.3 分配 & 3.4 假设\n（略，同标准模板）\n3.5 生命周期清单数据')

    # 动态生成表 3-2 到 3-11
    table_idx = 2
    for stage_name, records in user_records.items():
        doc.add_heading(stage_name, level=2)

        # 写入物料/能耗表
        mat_data = records["Material"]
        if mat_data:
            doc.add_paragraph(f'{stage_name}的清单数据，如表3-{table_idx}、3-{table_idx + 1}所示。')
            add_word_table(doc, f'表 3-{table_idx}：{stage_name}数据清单', ['名称', '消耗数量', '单位', '类型'],
                           mat_data)
            table_idx += 1

        # 写入运输表
        trans_data = records["Transport"]
        if trans_data:
            trans_formatted = [[row[0], "1", "道路运输", "-", row[1], "-"] for row in trans_data]  # 补齐模板字段
            add_word_table(doc, f'表 3-{table_idx}：{stage_name}运输数据清单',
                           ['运输物', '路段', '运输方式', '货物重量', '运输里程', '能源消耗'], trans_formatted)
            table_idx += 1

    # ---- Part 4: 评价结果 ----
    doc.add_heading('4. 产品碳足迹评价结果', level=1)
    doc.add_paragraph(
        f'通过计算，功能单位（1个汽车动力电池）的全生命周期碳足迹为 {total_carbon:,.6f} kgCO2e，单位产品排放量为 {total_carbon:,.6f} kgCO2e/个，碳足迹的整体情况，如表4-1所示。')

    res_data = []
    for stage, em in results.items():
        pct = f"{(em / total_carbon * 100):.4f}" if total_carbon > 0 else "0.0000"
        res_data.append([stage.split(" ")[1], f"{em:,.6f}", pct])
    res_data.append(['总计', f"{total_carbon:,.6f}", '100'])
    add_word_table(doc, '表4-1：产品碳足迹评价结果', ['生命周期阶段', '排放量（kgCO2e）', '占比（%）'], res_data)

    # ---- Part 5: 不确定性分析 (按要求不做改动，静态输出) ----
    doc.add_heading('5. 不确定性分析', level=1)
    doc.add_paragraph(
        '5.1 不确定性分析方法\n产品碳足迹评价的不确定性分析，采用定性分析法，包括活动水平数据质量评级、排放因子的质量评级。')
    add_word_table(doc, '表5-1：活动水平数据质量评级', ['质量等级', '描述', '分值'],
                   [['好', '量测值...', '5'], ['较好', '计算值...', '3'], ['一般', '理论值...', '2'],
                    ['差', '参考文献...', '1']])

    # ---- Part 6: 结论 (动态数值融合模板文本) ----
    doc.add_heading('6. 结论', level=1)
    pcts = [(em / total_carbon * 100) if total_carbon > 0 else 0 for em in results.values()]
    doc.add_paragraph(
        f'体验账号汽车动力电池的产品碳足迹评价，涵盖的时间范围是2026年01月01日至2026年12月31日。功能单位（1个汽车动力电池）的全生命周期碳足迹为 {total_carbon:,.6f} kgCO2e，其中原材料获取阶段、生产制造阶段、分销和储存阶段、产品使用阶段、废弃处置阶段的排放占比分别为：{pcts[0]:.2f}%、{pcts[1]:.2f}%、{pcts[2]:.2f}%、{pcts[3]:.2f}%、{pcts[4]:.2f}%。')

    # ---- 附录：静态输出 ----
    doc.add_heading('附录：排放因子选择表', level=1)
    doc.add_paragraph('（详见原始报告附录表，由系统后台数据库统一管理支持计算）')

    # 写入流
    bio = io.BytesIO()
    doc.save(bio)
    st.download_button(
        label="📥 点击下载完整 6 章节 Word 测算报告",
        data=bio.getvalue(),
        file_name="产品碳足迹评价报告_完整版.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )