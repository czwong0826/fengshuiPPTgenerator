import io
import json

import streamlit as st
from pptx import Presentation
from knowledge_base import ROOM_KNOWLEDGE_BASE

st.set_page_config(page_title="风水报告生成器", layout="wide")

st.title("🏠 风水报告生成器（原型）")
st.caption("在左侧选择需要分析的区域，然后在下方对应面板中填写细节，最后点击底部按钮生成 JSON 数据或 PPT 文件。")

# ---------- 1. 侧边栏多选框 ----------
AREA_OPTIONS = [
    "大门", "客厅", "饭厅", "干厨房", "湿厨房",
    "主卧", "未来儿子房", "未来女儿房", "卫生间",
]

with st.sidebar:
    st.header("选择区域")
    selected_areas = st.multiselect(
        "请选择需要分析的区域",
        options=AREA_OPTIONS,
        default=[],
    )
    st.markdown("---")
    st.caption(f"已选择 {len(selected_areas)} 个区域")

# ---------- 2 & 3. 动态生成折叠面板 ----------

report_data = {}

if not selected_areas:
    st.info("请先在左侧边栏选择至少一个区域，下方将自动生成对应的填写面板。")
else:
    st.subheader("区域详情填写")
    for area in selected_areas:
        with st.expander(f"📍 {area}", expanded=True):
            checked_items = []
            # 获取当前区域对应的选项，如果没有匹配到，则提供通用选项
            current_area_items = ROOM_KNOWLEDGE_BASE.get(area, ["保持整洁通风", "注意采光"])

            for idx, item in enumerate(current_area_items):
                checked = st.checkbox(item, key=f"{area}_checkbox_{idx}")
                if checked:
                    checked_items.append(item)

            note = st.text_area(
                "额外备注",
                key=f"{area}_note",
                placeholder=f"请输入针对「{area}」的额外说明...",
                height=100,
            )

            report_data[area] = {
                "已勾选项": checked_items,
                "备注": note,
            }

st.markdown("---")


# ---------- PPTX 生成函数 ----------
def build_pptx(data: dict) -> bytes:
    """遍历 report_data，为每个区域生成一页幻灯片，返回 PPTX 的字节流。"""
    prs = Presentation()
    # layout 1 = "Title and Content"（默认空白模板中，带标题+正文占位符的版式）
    title_and_content_layout = prs.slide_layouts[1]

    for area, content in data.items():
        checked_items = content.get("已勾选项", [])
        note = content.get("备注", "").strip()

        slide = prs.slides.add_slide(title_and_content_layout)
        slide.shapes.title.text = area

        body_placeholder = slide.placeholders[1]
        text_frame = body_placeholder.text_frame
        text_frame.clear()  # 清空默认占位文本

        bullets_written = False

        # 已勾选项逐条作为项目符号写入
        for item in checked_items:
            if not bullets_written:
                text_frame.text = item
                bullets_written = True
            else:
                p = text_frame.add_paragraph()
                p.text = item
            text_frame.paragraphs[-1].level = 0

        # 备注作为文本追加在下方（用低一级缩进区分）
        if note:
            if not bullets_written:
                text_frame.text = f"备注：{note}"
                bullets_written = True
            else:
                p = text_frame.add_paragraph()
                p.text = f"备注：{note}"
                p.level = 1

        # 如果既没有勾选项也没有备注，避免留空白正文
        if not bullets_written:
            text_frame.text = "（未填写具体内容）"

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------- 4. 生成 JSON 数据 / PPT 文件 ----------
col1, col2 = st.columns(2)

with col1:
    if st.button("生成 JSON 数据", type="primary", use_container_width=True):
        if not selected_areas:
            st.warning("尚未选择任何区域，无法生成报告。")
        else:
            final_output = {
                "选择的区域": selected_areas,
                "详情": report_data,
            }
            st.success("JSON 数据生成成功！")
            st.json(final_output)

            json_str = json.dumps(final_output, ensure_ascii=False, indent=2)
            st.download_button(
                label="下载 JSON 文件",
                data=json_str,
                file_name="fengshui_report.json",
                mime="application/json",
            )

with col2:
    if st.button("生成 PPT", type="secondary", use_container_width=True):
        if not selected_areas:
            st.warning("尚未选择任何区域，无法生成报告。")
        else:
            pptx_bytes = build_pptx(report_data)
            st.success("PPT 生成成功！")
            st.download_button(
                label="下载 PPTX 文件",
                data=pptx_bytes,
                file_name="fengshui_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )