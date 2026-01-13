import streamlit as st
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from io import BytesIO
import os

# ======================================================
# 主类
# ======================================================
class StreamlitCalculator:
    def __init__(self):
        if "stock_solutions" not in st.session_state:
            self.init_stock_solutions()
        if "molecular_weights" not in st.session_state:
            self.init_mw()
        if "excel_content" not in st.session_state:
            st.session_state.excel_content = None
        if "calculation_success" not in st.session_state:
            st.session_state.calculation_success = False

    # ------------------------
    # 初始化
    # ------------------------
    def init_stock_solutions(self):
        st.session_state.stock_solutions = {
            "Tris": {"concentration": 2.0, "unit": "M", "density": 1.0},
            "NaCl": {"concentration": 5.0, "unit": "M", "density": 1.0},
            "甘油": {"concentration": 100.0, "unit": "%", "density": 1.26},
            "DTT": {"concentration": 1.0, "unit": "M", "density": 1.0},
            "PBS": {"concentration": 10.0, "unit": "X", "density": 1.0},
            "CHAPS": {"concentration": 10.0, "unit": "%", "density": 1.0},
        }

    def init_mw(self):
        st.session_state.molecular_weights = {
            "Tris": 121.14,
            "NaCl": 58.44,
            "甘油": 92.09,
            "DTT": 154.25,
            "CHAPS": 614.88,
        }

    # ------------------------
    # 解析配方
    # ------------------------
    def parse_formula_string(self, formula_input):
        formula_input = re.sub(r'[，；、]', ',', formula_input)
        pattern = r'([\d\.]+)\s*([mMμu%Xx]*)\s*([a-zA-Z\u4e00-\u9fa5\-]+)'
        matches = re.findall(pattern, formula_input)

        components = {}
        for value, unit, name in matches:
            unit = unit.upper().replace("U", "μ")
            components[name] = {
                "target_concentration": float(value),
                "target_unit": unit if unit else "mM"
            }
        return components

    # ------------------------
    # 体积解析
    # ------------------------
    def parse_volume(self, text):
        text = text.strip().lower()
        m = re.match(r'([\d\.]+)\s*(l|ml|ul|μl)?', text)
        if not m:
            return None
        v = float(m.group(1))
        unit = m.group(2) or "ml"
        if unit == "l":
            return v * 1000
        if unit in ["ul", "μl"]:
            return v / 1000
        return v

    # ------------------------
    # 计算
    # ------------------------
    def calculate(self, components, total_ml):
        results = {"components": {}, "total": 0}

        for name, info in components.items():
            if name in st.session_state.stock_solutions:
                stock = st.session_state.stock_solutions[name]
                v = (info["target_concentration"] * total_ml) / stock["concentration"]
                results["components"][name] = {
                    "target": f'{info["target_concentration"]} {info["target_unit"]}',
                    "volume": v,
                    "mass": v * stock["density"]
                }
                results["total"] += v

            elif name in st.session_state.molecular_weights:
                mw = st.session_state.molecular_weights[name]
                mol = (info["target_concentration"] / 1000) * (total_ml / 1000)
                mass = mol * mw
                results["components"][name] = {
                    "target": f'{info["target_concentration"]} {info["target_unit"]}',
                    "volume": 0,
                    "mass": mass
                }

        water = total_ml - results["total"]
        results["components"]["水"] = {
            "target": "-",
            "volume": water,
            "mass": water
        }
        return results

    # ------------------------
    # 写入 Excel（模板版）
    # ------------------------
    def write_to_excel(self, formula_input, results, total_ml):
        try:
            # ⚠️ 模板必须和 app.py 在同一目录
            wb = load_workbook("template.xlsx")
            ws = wb.active

            ws["C5"] = datetime.now().strftime("%Y-%m-%d")
            ws["C6"] = formula_input
            ws["G6"] = f"{total_ml/1000:.2f} L"

            start_col = 3
            comps = [(k, v) for k, v in results["components"].items() if k != "水"]

            for i, (name, comp) in enumerate(comps):
                col = get_column_letter(start_col + i)
                ws[f"{col}8"] = name
                ws[f"{col}11"] = comp["target"]
                ws[f"{col}12"] = round(comp["mass"], 2) if comp["mass"] > 0 else "-"
                ws[f"{col}13"] = round(comp["volume"], 2) if comp["volume"] > 0 else "-"

            water_col = get_column_letter(start_col + len(comps))
            ws[f"{water_col}8"] = "水"
            ws[f"{water_col}12"] = round(results["components"]["水"]["mass"], 2)
            ws[f"{water_col}13"] = round(results["components"]["水"]["volume"], 2)

            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.session_state.excel_content = buffer
            return True

        except Exception as e:
            st.error(f"Excel 生成失败: {e}")
            return False

    # ------------------------
    # UI
    # ------------------------
    def run(self):
        st.title("🧪 试剂配方计算器（模板版）")

        formula = st.text_area(
            "配方输入",
            "20 mM Tris, 150 mM NaCl\n1 mM DTT",
            height=150
        )

        volume = st.text_input("目标体积", "1 L")

        if st.button("🚀 开始计算", type="primary"):
            total_ml = self.parse_volume(volume)
            if not total_ml:
                st.error("体积格式错误")
                return

            comps = self.parse_formula_string(formula)
            results = self.calculate(comps, total_ml)

            ok = self.write_to_excel(formula, results, total_ml)
            if ok:
                st.session_state.calculation_success = True
                st.success("计算完成，可下载 Excel")

        if st.session_state.calculation_success and st.session_state.excel_content:
            st.download_button(
                "📥 下载 Excel",
                st.session_state.excel_content,
                file_name="配方计算结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ======================================================
# 主入口
# ======================================================
if __name__ == "__main__":
    st.set_page_config(page_title="试剂配方计算器", page_icon="🧪")
    StreamlitCalculator().run()
