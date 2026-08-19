import io
import json
import uuid

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
st.caption("在左侧添加需要分析的区域（支持同类型多个实例，如厕所A/厕所B），上传平面图并点击打点，然后在下方填写细节，最后生成 JSON 数据或 PPT 文件。")

CUSTOM_TYPE = "自定义区域"
# 房间类型列表：知识库里的所有类型 + 一个"自定义区域"（无预设选项，只能写备注）
ROOM_TYPE_OPTIONS = list(ROOM_KNOWLEDGE_BASE.keys()) + [CUSTOM_TYPE]

# ---------- session_state 初始化 ----------
# instances: 已添加的区域实例列表，每项是 {"id": 唯一ID, "type": 房间类型, "name": 显示名称}
# area_coordinates: 打点坐标，key 是实例的唯一 ID（不是名字！同名/同类型实例也不会互相覆盖）
st.session_state.setdefault("instances", [])
st.session_state.setdefault("area_coordinates", {})
# 用递增计数器给"添加区域"表单里的控件生成新 key，强制它们在每次提交后
# 重新以默认值渲染（相当于手动实现清空效果，不依赖 clear_on_submit）
st.session_state.setdefault("add_form_version", 0)


def generate_default_name(room_type: str) -> str:
    """同类型自动编号：卫生间 1、卫生间 2……自定义区域默认叫"未命名区域 N"。"""
    base = "未命名区域" if room_type == CUSTOM_TYPE else room_type
    count = sum(1 for inst in st.session_state["instances"] if inst["type"] == room_type)
    return f"{base} {count + 1}"


# ---------- 1. 侧边栏：添加 / 管理区域实例 ----------
with st.sidebar:
    st.header("➕ 添加区域")

    form_version = st.session_state["add_form_version"]
    with st.form("add_area_form", clear_on_submit=True):
        selected_type = st.selectbox(
            "房间类型", options=ROOM_TYPE_OPTIONS, key=f"new_area_type_{form_version}"
        )
        custom_name = st.text_input(
            "显示名称（可选，留空自动编号）",
            placeholder="例如：厕所 A",
            key=f"new_area_name_{form_version}",
        )
        submitted = st.form_submit_button("添加区域", use_container_width=True)

        if submitted:
            final_name = custom_name.strip() or generate_default_name(selected_type)
            st.session_state["instances"].append(
                {
                    "id": uuid.uuid4().hex[:8],
                    "type": selected_type,
                    "name": final_name,
                }
            )
            st.session_state["add_form_version"] += 1
            st.toast(f"已添加「{final_name}」")
            st.rerun()

    st.markdown("---")
    st.subheader("已添加区域")

    if not st.session_state["instances"]:
        st.caption("尚未添加任何区域，请使用上方表单添加。")
    else:
        for inst in list(st.session_state["instances"]):
            col1, col2 = st.columns([4, 1])
            col1.write(f"📍 **{inst['name']}**　`{inst['type']}`")
            if col2.button("🗑️", key=f"del_{inst['id']}", help="删除该区域"):
                st.session_state["instances"] = [
                    i for i in st.session_state["instances"] if i["id"] != inst["id"]
                ]
                st.session_state["area_coordinates"].pop(inst["id"], None)
                st.rerun()

    st.markdown("---")
    st.caption(f"共 {len(st.session_state['instances'])} 个区域")

instances = st.session_state["instances"]

# ---------- 2. 平面图上传 + 点击打点 ----------
st.subheader("🗺️ 平面图打点标注")

uploaded_image = st.file_uploader("上传平面图（用于标注各区域位置）", type=["png", "jpg", "jpeg"])

floor_plan_bytes = None
img_width = None
img_height = None

if uploaded_image is not None:
    floor_plan_bytes = uploaded_image.getvalue()
    pil_image = Image.open(io.BytesIO(floor_plan_bytes))
    img_width, img_height = pil_image.size

    if not instances:
        st.info("请先在左侧添加至少一个区域，才能在图上为其打点。")
    else:
        # 下拉框展示"显示名称"，但实际选中值是实例的唯一 ID
        target_instance_id = st.selectbox(
            "选择要标注的区域（点击下方图片即可为该区域记录坐标）",
            options=[inst["id"] for inst in instances],
            format_func=lambda iid: next(
                (i["name"] for i in instances if i["id"] == iid), iid
            ),
            key="target_instance_for_click",
        )
        target_name = next(i["name"] for i in instances if i["id"] == target_instance_id)

        st.caption(f"图片原始尺寸：{img_width} × {img_height} px。点击图片，将为「{target_name}」记录打点位置。")

        coords = streamlit_image_coordinates(pil_image, key="floor_plan_coords")

        if coords is not None:
            # streamlit_image_coordinates 会在没有新点击时持续返回上一次的坐标，
            # 用签名比较判断这是不是一次"新"点击，避免每次 rerun 都覆盖当前选中的目标区域
            coord_signature = (coords["x"], coords["y"])
            if st.session_state.get("last_coord_signature") != coord_signature:
                st.session_state["last_coord_signature"] = coord_signature
                st.session_state["area_coordinates"][target_instance_id] = {
                    "x": coords["x"],
                    "y": coords["y"],
                    "width": img_width,
                    "height": img_height,
                }
                st.success(f"已记录「{target_name}」的坐标：({coords['x']}, {coords['y']})")

        if st.session_state["area_coordinates"]:
            with st.expander("📌 已记录的坐标一览", expanded=False):
                # 展示时把 ID 换成人类可读的名字
                readable = {
                    next((i["name"] for i in instances if i["id"] == iid), iid): coord
                    for iid, coord in st.session_state["area_coordinates"].items()
                }
                st.json(readable)
                if st.button("清空所有坐标"):
                    st.session_state["area_coordinates"] = {}
                    st.session_state.pop("last_coord_signature", None)
                    st.rerun()
else:
    st.info("上传平面图后，可在图上为每个已添加区域点击打点。")

st.markdown("---")


# ---------- 渲染单个知识库条目（普通 checkbox 或 主选项+子选项） ----------
def render_knowledge_item(instance_id: str, idx: int, item):
    """渲染一个条目。key 前缀一律使用实例的唯一 ID，
    避免同类型/同名的多个实例互相覆盖彼此的勾选状态。
    普通字符串直接渲染 checkbox；dict 类型（选择型条目）渲染 checkbox + 子选项
    （单选/多选），返回勾选后拼好的最终文案；未勾选或子选项未选完整时返回 None。
    """
    if isinstance(item, str):
        checked = st.checkbox(item, key=f"{instance_id}_checkbox_{idx}")
        return item if checked else None

    # 选择型条目
    checked = st.checkbox(item["label"], key=f"{instance_id}_mainchk_{idx}")
    if not checked:
        return None

    values = {}
    all_filled = True
    with st.container():
        for group_key, group_cfg in item["choices"].items():
            options = group_cfg["options"]
            if group_cfg["multi"]:
                selected = st.multiselect(
                    "　↳ 请选择具体内容",
                    options=options,
                    key=f"{instance_id}_subms_{idx}_{group_key}",
                )
                if not selected:
                    all_filled = False
                else:
                    values[group_key] = "、".join(selected)
            else:
                selected = st.radio(
                    "　↳ 请选择具体内容",
                    options=options,
                    key=f"{instance_id}_subradio_{idx}_{group_key}",
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


# ---------- 3. 动态生成折叠面板（遍历区域实例） ----------
# report_data 用实例 ID 做 key，避免同名/同类型区域互相覆盖
report_data = {}

if not instances:
    st.info("请先在左侧添加至少一个区域，下方将自动生成对应的填写面板。")
else:
    st.subheader("区域详情填写")
    for inst in instances:
        instance_id = inst["id"]
        area_type = inst["type"]
        area_name = inst["name"]

        with st.expander(f"📍 {area_name}", expanded=True):
            checked_items = []
            # 知识库里查不到该类型（自定义区域，或知识库以外的类型）就返回空列表，
            # 折叠面板里只剩备注框和打点信息，不报错
            current_area_items = ROOM_KNOWLEDGE_BASE.get(area_type, [])

            if not current_area_items:
                st.caption("（该区域无预设勾选项，可在下方直接填写备注）")

            for idx, item in enumerate(current_area_items):
                result_text = render_knowledge_item(instance_id, idx, item)
                if result_text:
                    checked_items.append(result_text)

            note = st.text_area(
                "额外备注",
                key=f"{instance_id}_note",
                placeholder=f"请输入针对「{area_name}」的额外说明...",
                height=100,
            )

            report_data[instance_id] = {
                "名称": area_name,
                "类型": area_type,
                "已勾选项": checked_items,
                "备注": note,
            }

            # 把该区域记录的打点坐标合并进 report_data
            if instance_id in st.session_state["area_coordinates"]:
                coord = st.session_state["area_coordinates"][instance_id]
                report_data[instance_id]["坐标"] = coord
                st.caption(f"📌 已标注坐标：({coord['x']}, {coord['y']})")

st.markdown("---")


# ---------- PPTX 生成函数 ----------
def build_pptx(
    area_reports: list,
    floor_plan_bytes: bytes = None,
    img_width: int = None,
    img_height: int = None,
) -> bytes:
    """遍历区域实例列表，为每个实例生成一页幻灯片：
    左侧放文本内容（勾选项 + 备注），右侧放平面图，
    如果该区域记录了打点坐标，就在图上对应位置画一个红色圆点。
    幻灯片标题使用『显示名称』，同类型的多个实例（如厕所A/厕所B）会各自独立成页。
    """
    prs = Presentation()
    title_and_content_layout = prs.slide_layouts[1]

    # 图片显示区域的边界框（幻灯片右侧），图片会按原始宽高比缩放后居中放入此框
    IMG_BOX_LEFT = Inches(5.3)
    IMG_BOX_TOP = Inches(1.6)
    IMG_BOX_MAX_WIDTH = Inches(4.2)
    IMG_BOX_MAX_HEIGHT = Inches(5.3)

    for report in area_reports:
        area_name = report.get("名称", "未命名区域")
        checked_items = report.get("已勾选项", [])
        note = report.get("备注", "").strip()
        coord = report.get("坐标")

        slide = prs.slides.add_slide(title_and_content_layout)
        slide.shapes.title.text = area_name

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


# ---------- 4. 生成 JSON 数据 / PPT 文件 ----------
col1, col2 = st.columns(2)

# 按实例添加顺序整理成列表（而不是用不透明的 UUID 做 key 展示给用户）
ordered_reports = [report_data[inst["id"]] for inst in instances if inst["id"] in report_data]

with col1:
    if st.button("生成 JSON 数据", type="primary", use_container_width=True):
        if not instances:
            st.warning("尚未添加任何区域，无法生成报告。")
        else:
            final_output = {"区域列表": ordered_reports}
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
        if not instances:
            st.warning("尚未添加任何区域，无法生成报告。")
        else:
            pptx_bytes = build_pptx(
                ordered_reports,
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