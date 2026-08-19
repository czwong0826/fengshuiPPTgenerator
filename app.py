import io
import json

import streamlit as st
from PIL import Image
from streamlit_image_coordinates import streamlit_image_coordinates

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from knowledge_base import ROOM_KNOWLEDGE_BASE

st.set_page_config(page_title="风水报告生成器", layout="wide")

st.title("🏠 风水报告生成器（原型）")
st.caption("在左侧选择需要分析的区域，上传平面图并点击打点，然后在下方对应面板中填写细节，最后点击底部按钮生成 JSON 数据或 PPT 文件。")

# ---------- 1. 侧边栏多选框 ----------
# 直接从知识库的 key 动态生成，避免"知识库加了新区域、侧边栏却忘记同步"的问题
AREA_OPTIONS = list(ROOM_KNOWLEDGE_BASE.keys())

with st.sidebar:
    st.header("选择区域")
    selected_areas = st.multiselect(
        "请选择需要分析的区域",
        options=AREA_OPTIONS,
        default=[],
    )
    st.markdown("---")
    st.caption(f"已选择 {len(selected_areas)} 个区域")

# ---------- 2. 平面图上传 + 点击打点 ----------
st.subheader("🗺️ 平面图打点标注")

uploaded_image = st.file_uploader("上传平面图（用于标注各区域位置）", type=["png", "jpg", "jpeg"])

floor_plan_bytes = None
img_width = None
img_height = None

# session_state 用于跨 rerun 持久化每个区域的打点坐标
# （report_data 每次都是在下方循环里重新构建的，坐标必须单独存起来再合并进去）
st.session_state.setdefault("area_coordinates", {})

if uploaded_image is not None:
    floor_plan_bytes = uploaded_image.getvalue()
    pil_image = Image.open(io.BytesIO(floor_plan_bytes))
    img_width, img_height = pil_image.size

    if not selected_areas:
        st.info("请先在左侧边栏至少选择一个区域，才能在图上为其打点。")
    else:
        target_area = st.selectbox(
            "选择要标注的区域（点击下方图片即可为该区域记录坐标）",
            options=selected_areas,
            key="target_area_for_click",
        )

        st.caption(f"图片原始尺寸：{img_width} × {img_height} px。点击图片，将为「{target_area}」记录打点位置。")

        coords = streamlit_image_coordinates(pil_image, key="floor_plan_coords")

        if coords is not None:
            # streamlit_image_coordinates 会在没有新点击时持续返回上一次的坐标，
            # 用签名比较判断这是不是一次“新”点击，避免每次 rerun 都覆盖当前选中的区域
            coord_signature = (coords["x"], coords["y"])
            if st.session_state.get("last_coord_signature") != coord_signature:
                st.session_state["last_coord_signature"] = coord_signature
                st.session_state["area_coordinates"][target_area] = {
                    "x": coords["x"],
                    "y": coords["y"],
                    "width": img_width,
                    "height": img_height,
                }
                st.success(f"已记录「{target_area}」的坐标：({coords['x']}, {coords['y']})")

        if st.session_state["area_coordinates"]:
            with st.expander("📌 已记录的坐标一览", expanded=False):
                st.json(st.session_state["area_coordinates"])
                if st.button("清空所有坐标"):
                    st.session_state["area_coordinates"] = {}
                    st.session_state.pop("last_coord_signature", None)
                    st.rerun()
else:
    st.info("上传平面图后，可在图上为每个已选区域点击打点。")

st.markdown("---")


# ---------- 渲染单个知识库条目（普通 checkbox 或 主选项+子选项） ----------
def render_knowledge_item(area: str, idx: int, item):
    """渲染一个条目。普通字符串直接渲染 checkbox；dict 类型（选择型条目）
    渲染 checkbox + 子选项（单选/多选），返回勾选后拼好的最终文案；
    未勾选或子选项未选完整时返回 None。
    """
    if isinstance(item, str):
        checked = st.checkbox(item, key=f"{area}_checkbox_{idx}")
        return item if checked else None

    # 选择型条目
    checked = st.checkbox(item["label"], key=f"{area}_mainchk_{idx}")
    if not checked:
        return None

    values = {}
    all_filled = True
    with st.container():
        for group_key, group_cfg in item["choices"].items():
            options = group_cfg["options"]
            if group_cfg["multi"]:
                selected = st.multiselect(
                    f"　↳ 请选择具体内容",
                    options=options,
                    key=f"{area}_subms_{idx}_{group_key}",
                )
                if not selected:
                    all_filled = False
                else:
                    values[group_key] = "、".join(selected)
            else:
                selected = st.radio(
                    f"　↳ 请选择具体内容",
                    options=options,
                    key=f"{area}_subradio_{idx}_{group_key}",
                    index=None,
                )
                if selected is None:
                    all_filled = False
                else:
                    values[group_key] = selected

    if not all_filled:
        st.caption("⚠️ 请完成上方子选项的选择，否则该项不会计入报告")
        return None

    return item["template"].format(**values)


# ---------- 3 & 4. 动态生成折叠面板 ----------
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
                result_text = render_knowledge_item(area, idx, item)
                if result_text:
                    checked_items.append(result_text)

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

            # 把该区域记录的打点坐标合并进 report_data
            if area in st.session_state["area_coordinates"]:
                coord = st.session_state["area_coordinates"][area]
                report_data[area]["坐标"] = coord
                st.caption(f"📌 已标注坐标：({coord['x']}, {coord['y']})")

st.markdown("---")


# ---------- PPTX 生成函数 ----------
def build_pptx(
    data: dict,
    floor_plan_bytes: bytes = None,
    img_width: int = None,
    img_height: int = None,
) -> bytes:
    """遍历 report_data，为每个区域生成一页幻灯片：
    左侧放文本内容（勾选项 + 备注），右侧放平面图，
    如果该区域记录了打点坐标，就在图上对应位置画一个红色圆点。
    """
    prs = Presentation()
    title_and_content_layout = prs.slide_layouts[1]

    # 图片显示区域的边界框（幻灯片右侧），图片会按原始宽高比缩放后居中放入此框
    IMG_BOX_LEFT = Inches(5.3)
    IMG_BOX_TOP = Inches(1.6)
    IMG_BOX_MAX_WIDTH = Inches(4.2)
    IMG_BOX_MAX_HEIGHT = Inches(5.3)

    for area, content in data.items():
        checked_items = content.get("已勾选项", [])
        note = content.get("备注", "").strip()
        coord = content.get("坐标")

        slide = prs.slides.add_slide(title_and_content_layout)
        slide.shapes.title.text = area

        # ---- 左侧：文本框（勾选项 + 备注） ----
        body_placeholder = slide.placeholders[1]
        # 收窄正文占位符宽度，让内容保持在幻灯片左半边，给右侧图片留出空间
        body_placeholder.left = Inches(0.5)
        body_placeholder.top = Inches(1.6)
        body_placeholder.width = Inches(4.5)
        body_placeholder.height = Inches(5.3)

        text_frame = body_placeholder.text_frame
        text_frame.word_wrap = True
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

        # ---- 右侧：平面图 + 打点标记 ----
        if floor_plan_bytes and img_width and img_height:
            aspect = img_width / img_height
            box_aspect = IMG_BOX_MAX_WIDTH / IMG_BOX_MAX_HEIGHT

            if aspect > box_aspect:
                # 图片相对更“宽”，以边界框宽度为约束
                final_width = IMG_BOX_MAX_WIDTH
                final_height = int(final_width / aspect)
            else:
                # 图片相对更“高”，以边界框高度为约束
                final_height = IMG_BOX_MAX_HEIGHT
                final_width = int(final_height * aspect)

            # 图片在边界框内居中放置
            pic_left = IMG_BOX_LEFT + int((IMG_BOX_MAX_WIDTH - final_width) / 2)
            pic_top = IMG_BOX_TOP + int((IMG_BOX_MAX_HEIGHT - final_height) / 2)

            # 每张幻灯片都需要一个独立的字节流（BytesIO 读取后指针会移动）
            image_stream = io.BytesIO(floor_plan_bytes)
            picture = slide.shapes.add_picture(
                image_stream,
                left=pic_left,
                top=pic_top,
                width=final_width,
                height=final_height,
            )

            # ---- 坐标几何换算：像素坐标 -> 比例 -> PPT 绝对位置 ----
            if coord:
                x_ratio = coord["x"] / coord["width"]
                y_ratio = coord["y"] / coord["height"]

                # 用图片在 PPT 中的实际 left/top/width/height（来自 picture 对象本身，
                # 避免自己心算产生偏差）换算出打点的绝对位置
                dot_center_x = picture.left + int(picture.width * x_ratio)
                dot_center_y = picture.top + int(picture.height * y_ratio)

                dot_size = Pt(14)
                marker = slide.shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    int(dot_center_x - dot_size / 2),
                    int(dot_center_y - dot_size / 2),
                    dot_size,
                    dot_size,
                )
                marker.fill.solid()
                marker.fill.fore_color.rgb = RGBColor(255, 0, 0)
                marker.line.color.rgb = RGBColor(255, 0, 0)
                marker.shadow.inherit = False

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------- 5. 生成 JSON 数据 / PPT 文件 ----------
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
            pptx_bytes = build_pptx(
                report_data,
                floor_plan_bytes=floor_plan_bytes,
                img_width=img_width,
                img_height=img_height,
            )
            st.success("PPT 生成成功！")
            st.download_button(
                label="下载 PPTX 文件",
                data=pptx_bytes,
                file_name="fengshui_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )