import io
import json

import streamlit as st
from pptx import Presentation

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
# ---------- 风水知识库 (根据房间动态匹配) ----------
ROOM_KNOWLEDGE_BASE = {
    "大门": [
        "大門建議常關",
        "木门为佳",
        "可以常挂红彩布和红色地毯",
        "可以放藍色地毯",
        "上方有壓梁情況，需拉平天花板或常關大门",
        "樓上廁所形成“淋頭水”，需化解"
    ],
    "客厅": [
        "顔色宜多用亮色为主",
        "忌用紅、粉紅、紫、橙色",
        "可有吊扇和吊灯",
        "可打开窗户通风纳气",
        "冷氣可在沙發上方"
    ],
    "饭厅": [
        "需常保明亮寬敞及整洁",
        "桌上不可放置雜物",
        "圓桌或長桌都適合",
        "餐桌不可透明玻璃"
    ],
    "干厨房": [
        "顔色以亮色爲主",
        "可做島臺設計",
        "切忌設立另外一個灶爐（包括電爐）"
    ],
    "湿厨房": [
        "灶需背靠实墙",
        "灶爐上廚忌用黑/灰/藍/紅色，建議白/青/木褐色",
        "建議使用電磁爐",
        "不能有吊扇",
        "垃圾桶不要靠近灶，建议有盖为佳"
    ],
    "卫生间": [
        "窗口需常開以便排煞",
        "盡量保持通風",
        "沒用時記得常關門",
        "需安裝抽風機把晦氣排出屋外"
    ],
    "主卧": [
        "床頭需背靠實墻",
        "常睡的房間不能有吊扇",
        "衣櫥材質不能選高反射或鏡面的",
        "需有床架与床頭櫃，床头柜不能高过床",
        "房間建議使用暖燈設計",
        "冷氣不可在床前方、房門 & 床頭上"
    ],
    "未来儿子房": [
        "床頭需背靠實墻",
        "常睡的房間不能有吊扇",
        "需有床架与床頭櫃，床头柜不能高过床",
        "房間建議使用暖燈設計"
    ],
    "未来女儿房": [
        "床頭需背靠實墻",
        "常睡的房間不能有吊扇",
        "需有床架与床頭櫃，床头柜不能高过床",
        "房間建議使用暖燈設計"
    ]
}

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