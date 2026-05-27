# -*- coding: utf-8 -*-
"""航空零件包装选型程序。"""

import math
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    from st_aggrid import AgGrid, GridOptionsBuilder
except ImportError:
    AgGrid = None
    GridOptionsBuilder = None


# Excel required columns.
BOX_COLS = {"model": "型号", "length": "内径长", "width": "内径宽", "height": "内径高"}
DIV_COLS = {"model": "型号", "length": "长度", "slots": "槽位数", "height": "高度"}

# Optional matching columns in carton sheet.
BOX_MATCH_COLUMNS = ["可用刀卡", "匹配刀卡", "指定刀卡", "刀卡型号", "可选刀卡"]
BOX_H_DIV_COLUMNS = ["横刀卡", "横刀", "横刀型号", "指定横刀", "匹配横刀"]
BOX_V_DIV_COLUMNS = ["竖刀卡", "竖刀", "竖刀型号", "指定竖刀", "匹配竖刀"]

# Optional divider cell-size limits.
DIV_MIN_SPACE_COLUMNS = ["最小空间", "最小格口", "最小间距"]
DIV_MAX_SPACE_COLUMNS = ["最大空间", "最大格口", "最大间距"]


def normalize_columns(df):
    """Trim Excel column names."""
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def require_columns(df, required_cols, label):
    """Validate required Excel columns."""
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"{label}缺少列: {', '.join(missing)}")


def first_existing_column(df, candidates):
    """Return the first candidate column that exists in a DataFrame."""
    for col in candidates:
        if col in df.columns:
            return col
    return None


def parse_model_list(value):
    """Parse model list such as '2,3,4' or '2，3，4'."""
    if pd.isna(value):
        return set()
    text = str(value).strip()
    if not text:
        return set()
    return {item.strip() for item in re.split(r"[,，;；/、\s]+", text) if item.strip()}


def model_matches(series, allowed_models):
    """Filter divider rows by allowed model set. Empty set means no limit."""
    if not allowed_models:
        return pd.Series([True] * len(series), index=series.index)
    return series.astype(str).str.strip().isin(allowed_models)


def get_optional_number(row, columns):
    """Read an optional numeric value from the first existing candidate column."""
    for col in columns:
        if col in row.index and pd.notna(row[col]):
            return float(row[col])
    return None


def cell_allowed_by_divider(cell_size, divider_row):
    """Check whether a cell size satisfies divider min/max space limits."""
    min_space = get_optional_number(divider_row, DIV_MIN_SPACE_COLUMNS)
    max_space = get_optional_number(divider_row, DIV_MAX_SPACE_COLUMNS)
    if min_space is not None and cell_size < min_space:
        return False
    if max_space is not None and cell_size > max_space:
        return False
    return True


def build_remark(capacity, target_qty, box_count, cell_utilization, order_box_utilization):
    """Build a readable result remark."""
    notes = []
    if box_count == 1:
        notes.append("单箱满足目标装量")
    else:
        notes.append(f"需{box_count}个箱子完成目标装量")

    if target_qty <= 1:
        notes.append("单件优先小箱")

    if cell_utilization >= 0.75:
        notes.append("格口紧凑")
    elif cell_utilization >= 0.55:
        notes.append("格口适中")
    else:
        notes.append("格口偏松，建议增加填充")

    if order_box_utilization < 0.08:
        notes.append("整单箱体空间浪费偏大")
    elif order_box_utilization >= 0.25:
        notes.append("整单箱体占用率较高")

    return "；".join(notes)


def possible_divider_counts(h_slots, v_slots):
    """Generate valid horizontal and vertical divider count pairs.

    Business rule:
    - horizontal divider count <= vertical divider slot count
    - vertical divider count <= horizontal divider slot count
    - if horizontal count is 0, vertical count must also be 0
    - once dividers are used, both horizontal and vertical counts must be > 0
    """
    max_h_count = int(v_slots) if pd.notna(v_slots) else 0
    max_v_count = int(h_slots) if pd.notna(h_slots) else 0

    for h_count in range(0, max_h_count + 1):
        for v_count in range(0, max_v_count + 1):
            no_dividers = h_count == 0 and v_count == 0
            crossed_dividers = h_count > 0 and v_count > 0
            if no_dividers or crossed_dividers:
                yield h_count, v_count


def possible_layout_configs(h_slots, v_slots):
    """Generate divider layouts.

    There are two physical meanings:
    - grid: every divider is treated as an internal separator.
    - bordered grid: two edge dividers form support borders, extra dividers split inside.
    """
    for h_count, v_count in possible_divider_counts(h_slots, v_slots):
        if h_count == 0 and v_count == 0:
            yield "无刀卡", h_count, v_count, 1, 1
            continue

        yield "分格", h_count, v_count, v_count + 1, h_count + 1

        if h_count >= 2 and v_count >= 2:
            yield "边框分格", h_count, v_count, v_count - 1, h_count - 1


def part_orientations(length, width, height):
    """Generate unique part orientations mapped to carton length, width, height."""
    candidates = [
        (length, width, height, "长宽高不旋转"),
        (width, length, height, "长宽旋转"),
        (length, height, width, "宽转高度"),
        (height, length, width, "宽转高度且底面旋转"),
        (width, height, length, "长转高度"),
        (height, width, length, "长转高度且底面旋转"),
    ]
    seen = set()
    for pl, pw, ph, note in candidates:
        key = (pl, pw, ph)
        if key not in seen:
            seen.add(key)
            yield pl, pw, ph, note


def find_best_packaging_logic(part_dim, target_qty, boxes_df, divs_df, t=6):
    """Find packaging options for a part and target order quantity."""
    p_l, p_w, p_h = part_dim
    boxes_df = normalize_columns(boxes_df)
    divs_df = normalize_columns(divs_df)
    require_columns(boxes_df, BOX_COLS.values(), "纸箱库")
    require_columns(divs_df, DIV_COLS.values(), "刀卡库")

    match_col = first_existing_column(boxes_df, BOX_MATCH_COLUMNS)
    h_match_col = first_existing_column(boxes_df, BOX_H_DIV_COLUMNS)
    v_match_col = first_existing_column(boxes_df, BOX_V_DIV_COLUMNS)
    results = []

    for _, box in boxes_df.iterrows():
        try:
            b_l = float(box[BOX_COLS["length"]])
            b_w = float(box[BOX_COLS["width"]])
            b_h = float(box[BOX_COLS["height"]])
            box_model = str(box[BOX_COLS["model"]]).strip()

            # Prefer explicit horizontal/vertical divider candidates.
            if h_match_col and v_match_col:
                h_allowed_models = parse_model_list(box[h_match_col])
                v_allowed_models = parse_model_list(box[v_match_col])
                if not h_allowed_models or not v_allowed_models:
                    continue
                h_divs = divs_df[model_matches(divs_df[DIV_COLS["model"]], h_allowed_models)]
                v_divs = divs_df[model_matches(divs_df[DIV_COLS["model"]], v_allowed_models)]
                match_label = f"横:{','.join(sorted(h_allowed_models))} / 竖:{','.join(sorted(v_allowed_models))}"
            else:
                allowed_models = parse_model_list(box[match_col]) if match_col else set()
                h_divs = divs_df[model_matches(divs_df[DIV_COLS["model"]], allowed_models)]
                v_divs = h_divs
                match_label = ",".join(sorted(allowed_models)) if allowed_models else "不限"

            if h_divs.empty or v_divs.empty:
                continue

            # Try all three-dimensional part orientations.
            for pl, pw, ph, rotate_note in part_orientations(p_l, p_w, p_h):
                if pl > b_l or pw > b_w or ph > b_h:
                    continue

                for _, v_div in v_divs.iterrows():
                    for _, h_div in h_divs.iterrows():
                        v_height = float(v_div[DIV_COLS["height"]])
                        h_height = float(h_div[DIV_COLS["height"]])
                        if not math.isclose(v_height, h_height):
                            continue

                        div_height = v_height
                        if div_height > b_h:
                            continue

                        v_len = float(v_div[DIV_COLS["length"]])
                        h_len = float(h_div[DIV_COLS["length"]])
                        if v_len > b_w or h_len > b_l:
                            continue

                        max_layer_count = max(1, int((b_h - t) // (div_height + t)))
                        h_slots = float(h_div[DIV_COLS["slots"]])
                        v_slots = float(v_div[DIV_COLS["slots"]])

                        for layout_mode, h_count_per_layer, v_count_per_layer, n, m in possible_layout_configs(h_slots, v_slots):
                            has_support = layout_mode != "无刀卡"
                            if target_qty > 1 and not has_support:
                                continue
                            k = max_layer_count if has_support else 1

                            if layout_mode == "边框分格":
                                cell_l = (b_l - v_count_per_layer * t) / n
                                cell_w = (b_w - h_count_per_layer * t) / m
                            else:
                                cell_l = (b_l - v_count_per_layer * t) / n
                                cell_w = (b_w - h_count_per_layer * t) / m

                            if cell_l < pl or not cell_allowed_by_divider(cell_l, v_div):
                                continue

                            if cell_w < pw or not cell_allowed_by_divider(cell_w, h_div):
                                continue

                            cell_h = div_height
                            if cell_h < ph:
                                continue

                            part_count_l = n
                            part_count_w = m

                            capacity = int(part_count_l * part_count_w * k)
                            if capacity <= 0:
                                continue
                            if capacity > 1 and not has_support:
                                continue

                            cell_volume = cell_l * cell_w * cell_h
                            part_volume = p_l * p_w * p_h
                            box_volume = b_l * b_w * b_h
                            box_count = math.ceil(target_qty / capacity)
                            total_package_volume = box_volume * box_count

                            cell_utilization = part_volume / cell_volume
                            box_utilization = part_volume * capacity / box_volume
                            order_box_utilization = part_volume * target_qty / total_package_volume
                            dim_closeness = (
                                min(pl / cell_l, 1) *
                                min(pw / cell_w, 1) *
                                min(ph / cell_h, 1)
                            )

                            if target_qty <= 1:
                                score = (
                                    order_box_utilization * 0.75 +
                                    cell_utilization * 0.15 +
                                    dim_closeness * 0.10
                                )
                            else:
                                score = (
                                    order_box_utilization * 0.70 +
                                    cell_utilization * 0.15 +
                                    dim_closeness * 0.10 +
                                    box_utilization * 0.05
                                )

                            remark = build_remark(
                                capacity,
                                target_qty,
                                box_count,
                                cell_utilization,
                                order_box_utilization,
                            )

                            results.append({
                                "推荐纸箱": box_model,
                                "可用刀卡限制": match_label,
                                "结构方式": layout_mode,
                                "排布方式": f"{part_count_l}x{part_count_w}x{k}",
                                "单箱容量": capacity,
                                "建议箱数": box_count,
                                "格口长": cell_l,
                                "格口宽": cell_w,
                                "格口高": cell_h,
                                "箱体体积": box_volume,
                                "总包装体积": total_package_volume,
                                "箱体利用率": box_utilization,
                                "整单箱体利用率": order_box_utilization,
                                "格口利用率": cell_utilization,
                                "综合评分": score,
                                "横刀型号": h_div[DIV_COLS["model"]],
                                "横刀总数": h_count_per_layer * k,
                                "竖刀型号": v_div[DIV_COLS["model"]],
                                "竖刀总数": v_count_per_layer * k,
                                "备注": remark,
                                "raw": {
                                    "box": (b_l, b_w, b_h),
                                    "part": (pl, pw, ph),
                                    "layout": (part_count_l, part_count_w, k),
                                    "cell": (cell_l, cell_w, cell_h),
                                    "tight_gap": t,
                                    "div_height": div_height,
                                    "rotate_note": rotate_note,
                                    "layout_mode": layout_mode,
                                    "h_count_per_layer": h_count_per_layer,
                                    "v_count_per_layer": v_count_per_layer,
                                },
                            })
        except Exception:
            continue

    if not results:
        return pd.DataFrame()

    df_res = pd.DataFrame(results)
    if target_qty <= 1:
        return df_res.sort_values(
            by=["整单箱体利用率", "总包装体积", "综合评分"],
            ascending=[False, True, False],
        )

    return df_res.sort_values(
        by=["整单箱体利用率", "综合评分", "总包装体积"],
        ascending=[False, False, True],
    )


def _add_box_mesh(fig, x, y, z, color, opacity, name):
    """Add a cuboid mesh to a Plotly 3D figure."""
    fig.add_trace(go.Mesh3d(
        x=x,
        y=y,
        z=z,
        i=[7, 0, 0, 0, 4, 4, 6, 6, 4, 0, 3, 2],
        j=[3, 4, 1, 2, 5, 6, 5, 2, 0, 1, 6, 3],
        k=[0, 7, 2, 3, 6, 7, 1, 1, 5, 5, 7, 6],
        color=color,
        opacity=opacity,
        flatshading=True,
        name=name,
    ))


def draw_3d_layout(box_dim, part_dim_tuple, layout, vh, t=6, layout_mode="分格", h_count=0, v_count=0):
    """Draw a carton and divider grid in 3D."""
    b_l, b_w, b_h = box_dim
    pl, pw, _ = part_dim_tuple
    n, m, k = [int(x) for x in layout]

    fig = go.Figure()
    _add_box_mesh(
        fig,
        x=[0, b_l, b_l, 0, 0, b_l, b_l, 0],
        y=[0, 0, b_w, b_w, 0, 0, b_w, b_w],
        z=[0, 0, 0, 0, b_h, b_h, b_h, b_h],
        color="gray",
        opacity=0.04,
        name="纸箱",
    )

    div_t = 4
    total_gap_l = b_l - (n * pl + ((n - 1) if n > 1 else 0) * t)
    total_gap_w = b_w - (m * pw + ((m - 1) if m > 1 else 0) * t)
    side_gap_l = max(0, total_gap_l / 2)
    side_gap_w = max(0, total_gap_w / 2)

    for layer in range(k):
        z_bot = t + layer * (vh + t)
        z_top = z_bot + vh

        if layout_mode == "边框分格":
            h_positions = []
            if h_count >= 1:
                h_positions.append(t / 2)
            if h_count >= 2:
                h_positions.append(b_w - t / 2)
            if h_count > 2:
                for i in range(1, h_count - 1):
                    h_positions.append(i * b_w / (h_count - 1))
        else:
            h_positions = [side_gap_w + i * (pw + t) - t / 2 for i in range(1, m)]

        for y_pos in h_positions:
            _add_box_mesh(
                fig,
                x=[0, b_l, b_l, 0, 0, b_l, b_l, 0],
                y=[
                    y_pos - div_t / 2, y_pos - div_t / 2,
                    y_pos + div_t / 2, y_pos + div_t / 2,
                    y_pos - div_t / 2, y_pos - div_t / 2,
                    y_pos + div_t / 2, y_pos + div_t / 2,
                ],
                z=[z_bot, z_bot, z_bot, z_bot, z_top, z_top, z_top, z_top],
                color="gold",
                opacity=1.0,
                name="横刀板",
            )

        if layout_mode == "边框分格":
            v_positions = []
            if v_count >= 1:
                v_positions.append(t / 2)
            if v_count >= 2:
                v_positions.append(b_l - t / 2)
            if v_count > 2:
                for j in range(1, v_count - 1):
                    v_positions.append(j * b_l / (v_count - 1))
        else:
            v_positions = [side_gap_l + j * (pl + t) - t / 2 for j in range(1, n)]

        for x_pos in v_positions:
            _add_box_mesh(
                fig,
                x=[
                    x_pos - div_t / 2, x_pos - div_t / 2,
                    x_pos + div_t / 2, x_pos + div_t / 2,
                    x_pos - div_t / 2, x_pos - div_t / 2,
                    x_pos + div_t / 2, x_pos + div_t / 2,
                ],
                y=[0, b_w, b_w, 0, 0, b_w, b_w, 0],
                z=[z_bot, z_bot, z_bot, z_bot, z_top, z_top, z_top, z_top],
                color="saddlebrown",
                opacity=1.0,
                name="竖刀板",
            )

    fig.update_layout(
        scene=dict(
            aspectmode="data",
            xaxis_title="长 (mm)",
            yaxis_title="宽 (mm)",
            zaxis_title="高 (mm)",
            xaxis=dict(gridcolor="lightgray"),
            yaxis=dict(gridcolor="lightgray"),
            zaxis=dict(gridcolor="lightgray"),
            bgcolor="white",
        ),
        margin=dict(l=0, r=0, b=0, t=0),
        scene_camera=dict(eye=dict(x=1.5, y=1.5, z=1.2)),
    )
    return fig


def render_app():
    """Streamlit UI."""
    st.set_page_config(page_title="航空包装决策系统", layout="wide")
    st.title("航空零件包装选型 & 3D 模拟")

    if "res_data" not in st.session_state:
        st.session_state.res_data = None

    with st.sidebar:
        f_box = st.file_uploader("纸箱库", type=["xlsx"])
        f_div = st.file_uploader("刀卡库", type=["xlsx"])

        with st.form("calc_form"):
            l = st.number_input("零件长", 1, 5000, 150)
            w = st.number_input("零件宽", 1, 5000, 100)
            h = st.number_input("零件高", 1, 2000, 50)
            tq = st.number_input("目标装量", 1, 1000, 12)
            submit = st.form_submit_button("开始匹配方案", use_container_width=True)

    if not f_box or not f_div:
        st.info("请先在左侧上传纸箱库和刀卡库 Excel 文件。")
        st.caption("纸箱库建议增加两列：横刀卡、竖刀卡，例如横刀卡=2,3，竖刀卡=8,9。刀卡库可增加：最小空间、最大空间。")
        return

    if submit:
        try:
            st.session_state.res_data = find_best_packaging_logic(
                (l, w, h),
                tq,
                pd.read_excel(f_box),
                pd.read_excel(f_div),
            )
        except Exception as exc:
            st.error(str(exc))
            return

    res = st.session_state.res_data
    if res is None:
        st.info("输入零件尺寸和目标装量后，点击开始匹配方案。")
        return

    if res.empty:
        st.warning("未发现满足条件的方案。请检查目标装量、纸箱指定刀卡、刀卡空间范围和零件尺寸。")
        return

    disp = res.drop(columns=["raw"]).copy()
    for col in ["格口长", "格口宽", "格口高"]:
        disp[col] = disp[col].map(lambda x: f"{x:.1f}")
    for col in ["箱体体积", "总包装体积"]:
        disp[col] = disp[col].map(lambda x: f"{x:,.0f}")
    for col in ["箱体利用率", "整单箱体利用率", "格口利用率"]:
        disp[col] = disp[col].map(lambda x: f"{x:.2%}")
    disp["综合评分"] = disp["综合评分"].map(lambda x: f"{x:.3f}")

    st.subheader("包装方案优选表")
    if AgGrid and GridOptionsBuilder:
        grid_builder = GridOptionsBuilder.from_dataframe(disp)
        grid_builder.configure_default_column(
            filter="agSetColumnFilter",
            sortable=True,
            resizable=True,
            floatingFilter=False,
            minWidth=120,
            width=140,
            suppressMenu=False,
            menuTabs=["filterMenuTab", "generalMenuTab", "columnsMenuTab"],
            filterParams={
                "buttons": ["apply", "reset"],
                "closeOnApply": True,
            },
        )
        for narrow_col in ["单箱容量", "建议箱数", "格口长", "格口宽", "格口高", "综合评分"]:
            if narrow_col in disp.columns:
                grid_builder.configure_column(narrow_col, minWidth=100, width=110)
        for wide_col in ["推荐纸箱", "可用刀卡限制", "备注"]:
            if wide_col in disp.columns:
                grid_builder.configure_column(wide_col, minWidth=180, width=220)
        grid_builder.configure_grid_options(
            domLayout="normal",
            enableCellTextSelection=True,
            suppressHorizontalScroll=False,
            alwaysShowHorizontalScroll=True,
        )
        AgGrid(
            disp,
            gridOptions=grid_builder.build(),
            fit_columns_on_grid_load=False,
            height=420,
            theme="streamlit",
            allow_unsafe_jscode=True,
            enable_enterprise_modules=True,
        )
    else:
        st.caption("安装 streamlit-aggrid 后，表头会显示筛选框：pip install streamlit-aggrid")
        st.dataframe(disp, use_container_width=True)

    st.divider()
    c1, c2 = st.columns([1, 1.2])

    with c1:
        selected_idx = st.selectbox(
            "选择方案查看 BOM 和 3D：",
            res.index.tolist(),
            format_func=lambda idx: f"{res.loc[idx, '推荐纸箱']} - {res.loc[idx, '排布方式']}",
        )
        item = res.loc[selected_idx]
        box_count = int(item["建议箱数"])
        pad_count = int(item["raw"]["layout"][2] + 1)

        st.info(f"""
### 领料 BOM 清单
- **纸箱型号**：{item["推荐纸箱"]}
- **建议箱数**：{box_count} 个
- **结构方式**：{item["结构方式"]}
- **允许刀卡**：{item["可用刀卡限制"]}
- **横刀型号**：{item["横刀型号"]}，单箱 **{item["横刀总数"]}** 把，合计 **{item["横刀总数"] * box_count}** 把
- **竖刀型号**：{item["竖刀型号"]}，单箱 **{item["竖刀总数"]}** 把，合计 **{item["竖刀总数"] * box_count}** 把
- **格口空间**：{item["格口长"]:.1f} x {item["格口宽"]:.1f} x {item["格口高"]:.1f} mm
- **辅助材料**：平垫板单箱 **{pad_count}** 张，合计 **{pad_count * box_count}** 张
- **方向**：{item["raw"]["rotate_note"]}
- **结果**：{item["备注"]}
""")

    with c2:
        st.plotly_chart(
            draw_3d_layout(
                item["raw"]["box"],
                item["raw"]["part"],
                item["raw"]["layout"],
                item["raw"]["div_height"],
                layout_mode=item["raw"].get("layout_mode", "分格"),
                h_count=item["raw"].get("h_count_per_layer", 0),
                v_count=item["raw"].get("v_count_per_layer", 0),
            ),
            use_container_width=True,
        )


if __name__ == "__main__":
    render_app()
