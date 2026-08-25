"""
Инженерное приложение для раздельного и сводного расчета потребления чистого 100% NaOH
в мокрых скрубберах термических установок по утилизации:
1) Жидких отходов (КТОЖС, 1,5 м³/ч)
2) Твердых бытовых и промышленных отходов (ТБО, 170 кг/ч)
с расчетом удельного расхода чистого реагента на 1 кг отхода.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from pathlib import Path
import os
import io

from calculations import (
    calculate_liquid_waste_pollutants,
    calculate_tbo_pollutants,
    calculate_liquid_installation_naoh,
    calculate_tbo_installation_naoh,
    calculate_combined_installations_naoh,
    calculate_naoh_and_compliance,
    LIQUID_WASTE_DATASETS,
    TBO_WASTE_GROUPS,
    TBO_PRESETS,
    M_HCl, M_NaOH, M_SO2, M_Cl, M_S
)
from export_report import generate_word_report, generate_pdf_report, generate_excel_report

st.set_page_config(
    page_title="Расчет чистого 100% NaOH для скрубберов",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Стили
st.markdown("""
<style>
    .main-header {
        font-size: 2.0rem;
        font-weight: 700;
        color: #102C57;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4A5568;
        margin-bottom: 0.8rem;
    }
    .badge-info {
        background-color: #EBF8FF;
        border-left: 4px solid #3182CE;
        padding: 8px 12px;
        border-radius: 4px;
        margin-bottom: 12px;
        font-size: 0.9rem;
    }
    .badge-spec {
        background-color: #F0FFF4;
        border-left: 4px solid #38A169;
        padding: 10px 14px;
        border-radius: 6px;
        margin-bottom: 12px;
        font-size: 0.95rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        background-color: #F1F5F9;
        border-radius: 6px 6px 0px 0px;
        padding: 8px 16px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #102C57 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Заголовок
st.markdown('<div class="main-header">🧪 Расчет расхода чистого 100% NaOH в скрубберах газоочистки</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Раздельные расчеты для Установки жидких отходов (1,5 м³/ч) и Установки ТБО (170 кг/ч)</div>', unsafe_allow_html=True)

# ==============================================================================
# SIDEBAR: РЕЖИМ РАБОТЫ, СМЕНЫ И ТЕХНОЛОГИЧЕСКИЕ ПАРАМЕТРЫ
# ==============================================================================
st.sidebar.subheader("💨 Дымовые газы установок")
flue_gas_flow_liq = st.sidebar.number_input(
    "Расход газов (Жидкие отходы), нм³/ч",
    min_value=50.0, max_value=50000.0, value=2500.0, step=100.0,
    key="fg_liq_in",
    help="Объем сухих дымовых газов от Установки утилизации жидких отходов (1,5 м³/ч)"
)

flue_gas_flow_tbo = st.sidebar.number_input(
    "Расход газов (ТБО 170 кг/ч), нм³/ч",
    min_value=50.0, max_value=50000.0, value=800.0, step=50.0,
    key="fg_tbo_in",
    help="Объем сухих дымовых газов от Установки утилизации ТБО (170 кг/ч)"
)

flue_gas_flow = flue_gas_flow_liq + flue_gas_flow_tbo
st.sidebar.caption(f"ℹ️ Суммарный объем газов: **{flue_gas_flow:.0f} нм³/ч** ({flue_gas_flow_liq:.0f} + {flue_gas_flow_tbo:.0f})")

with st.sidebar.expander("⏱️ График эксплуатации и простои", expanded=True):
    hours_per_day = st.number_input(
        "Длительность смены в сутки, ч/сут",
        min_value=1.0, max_value=24.0, value=24.0, step=1.0,
        help="Количество рабочих часов установки в сутки (24 ч — непрерывно, 12 ч — 1 смена, 8 ч — дневная смена)"
    )
    
    operating_days_year = st.number_input(
        "Рабочих дней в году (с учетом ППР), дн/год",
        min_value=1.0, max_value=365.0, value=365.0, step=1.0,
        help="Количество рабочих дней в году с поправкой на планово-предупредительный ремонт и простои (например, 365, 330 или 300 дней)"
    )
    
    annual_hours = hours_per_day * operating_days_year
    st.info(f"🕒 Годовой фонд времени: **{annual_hours:.0f} ч/год** ({operating_days_year:.0f} дн × {hours_per_day:.0f} ч)")

with st.sidebar.expander("🏭 Параметры скруббера", expanded=True):
    eta_scrubber = st.slider(
        "Паспортная эффективность скруббера (η)",
        min_value=0.85, max_value=0.99, value=0.95, step=0.01,
        help="Паспортная/номинальная степень очистки скруббера (проверяется на достаточность для ПДК ИТС 9-2020)"
    )
    
    k_excess = st.slider(
        "Коэффициент избытка NaOH (k_изб)",
        min_value=1.00, max_value=1.40, value=1.15, step=0.05,
        help="Технологический запас щелочи для поддержания pH скрубберной жидкости"
    )

with st.sidebar.expander("🔥 Конверсия при 1100 °C", expanded=True):
    st.caption("Часть Cl и S связывается в золе (NaCl, KCl, CaSO₄)")
    k_conv_cl_liq = st.slider("Доля Cl → HCl (жидкие отходы)", 0.0, 1.0, 0.95, 0.01, help="Степень перехода Cl в HCl для жидких отходов (при распылении в факел)")
    k_conv_s_liq = st.slider("Доля S → SO₂ (жидкие отходы)", 0.0, 1.0, 0.85, 0.01, help="Степень перехода S в SO₂ для жидких отходов (с учетом связывания в золе)")
    k_conv_cl_tbo = st.slider("Доля Cl → HCl (ТБО)", 0.0, 1.0, 0.85, 0.01, help="Степень перехода Cl в HCl для ТБО (с учетом удержания NaCl/KCl в золе)")
    k_conv_s_tbo = st.slider("Доля S → SO₂ (ТБО)", 0.0, 1.0, 0.80, 0.01, help="Степень перехода S в SO₂ для ТБО (с учетом связывания CaSO₄ в золе)")


# Предварительная оценка общего потока отходов для пропорционального распределения газов
q_liq_init = st.session_state.get('liq_q_flow', 1.5)
m_tbo_init = st.session_state.get('tbo_m_input', 170.0)
total_feed_mass_est = (q_liq_init * 1000.0) + m_tbo_init

# ==============================================================================
# ОСНОВНЫЕ ВКЛАДКИ
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "💧 1. Установка: Жидкие отходы (1,5 м³/ч)",
    "🗑️ 2. Установка: ТБО (170 кг/ч)",
    "📊 3. Сводная ведомость и Сравнение",
    "📄 4. Пояснительная записка и Экспорт"
])

# ------------------------------------------------------------------------------
# ВКЛАДКА 1: УСТАНОВКА УТИЛИЗАЦИИ ЖИДКИХ ОТХОДОВ
# ------------------------------------------------------------------------------
with tab1:
    st.subheader("💧 Установка №1: Утилизация жидких отходов (КТОЖС)")
    st.markdown("Исходные концентрации активных кислых компонентов ($\text{Cl}^-$ и $\text{SO}_4^{2-}$) по **СП 320.1325800.2017 (Таблица Г.1)**:")
    
    col_q, col_active_set = st.columns([1, 2])
    with col_q:
        q_liq = st.number_input(
            "Расход жидких стоков (Q_liq), м³/ч",
            min_value=0.1, max_value=20.0, value=1.5, step=0.1,
            key="liq_q_flow"
        )
    with col_active_set:
        active_set_key = st.radio(
            "Расчетный набор данных по фильтрату полигона:",
            options=["young", "old", "custom"],
            format_func=lambda x: {
                "young": "📌 Набор 1: «Молодой полигон» (кислая фаза) — Пиковые нагрузки",
                "old": "📌 Набор 2: «Старый полигон» (метаногенная фаза) — Стабилизированный режим",
                "custom": "✏️ Пользовательские концентрации"
            }[x],
            horizontal=False,
            key="liq_set_select"
        )
        
    st.markdown("---")
    
    c_set1, c_set2 = st.columns(2)
    with c_set1:
        st.markdown("##### 🧪 Набор 1: «Молодой полигон» (кислая фаза)")
        mode_s1 = st.selectbox("Нагрузка (Набор 1):", ["Максимум (проектный)", "Среднее", "Минимум", "Свой ввод"], key="s1_mode")
        ds1 = LIQUID_WASTE_DATASETS["young"]
        val_cl1 = ds1["cl_range"][2] if mode_s1 == "Максимум (проектный)" else (ds1["cl_range"][1] if mode_s1 == "Среднее" else (ds1["cl_range"][0] if mode_s1 == "Минимум" else 5000.0))
        val_so41 = ds1["so4_range"][2] if mode_s1 == "Максимум (проектный)" else (ds1["so4_range"][1] if mode_s1 == "Среднее" else (ds1["so4_range"][0] if mode_s1 == "Минимум" else 1500.0))
        
        c_cl1 = st.number_input("Cl⁻ (Набор 1), мг/дм³", value=float(val_cl1), step=100.0, key="c_cl1")
        c_so41 = st.number_input("SO₄²⁻ (Набор 1), мг/дм³", value=float(val_so41), step=50.0, key="c_so41")
        res_liq1_feed = calculate_liquid_waste_pollutants(q_liq, c_cl1, c_so41, k_conv_cl_liq, k_conv_s_liq, dataset_name="Набор 1 (Молодой полигон)")

    with c_set2:
        st.markdown("##### 🧪 Набор 2: «Старый полигон» (метаногенная фаза)")
        mode_s2 = st.selectbox("Нагрузка (Набор 2):", ["Максимум (проектный)", "Среднее", "Минимум", "Свой ввод"], key="s2_mode")
        ds2 = LIQUID_WASTE_DATASETS["old"]
        val_cl2 = ds2["cl_range"][2] if mode_s2 == "Максимум (проектный)" else (ds2["cl_range"][1] if mode_s2 == "Среднее" else (ds2["cl_range"][0] if mode_s2 == "Минимум" else 2500.0))
        val_so42 = ds2["so4_range"][2] if mode_s2 == "Максимум (проектный)" else (ds2["so4_range"][1] if mode_s2 == "Среднее" else (ds2["so4_range"][0] if mode_s2 == "Минимум" else 400.0))
        
        c_cl2 = st.number_input("Cl⁻ (Набор 2), мг/дм³", value=float(val_cl2), step=100.0, key="c_cl2")
        c_so42 = st.number_input("SO₄²⁻ (Набор 2), мг/дм³", value=float(val_so42), step=50.0, key="c_so42")
        res_liq2_feed = calculate_liquid_waste_pollutants(q_liq, c_cl2, c_so42, k_conv_cl_liq, k_conv_s_liq, dataset_name="Набор 2 (Старый полигон)")

    if active_set_key == "young":
        active_liq_feed = res_liq1_feed
        active_c_cl = c_cl1
        active_c_so4 = c_so41
        active_liq_title = "Набор 1: Молодой полигон (кислая фаза)"
    elif active_set_key == "old":
        active_liq_feed = res_liq2_feed
        active_c_cl = c_cl2
        active_c_so4 = c_so42
        active_liq_title = "Набор 2: Старый полигон (метаногенная фаза)"
    else:
        st.markdown("##### ✏️ Пользовательские концентрации стоков:")
        col_cu1, col_cu2 = st.columns(2)
        with col_cu1:
            active_c_cl = st.number_input("Cl⁻, мг/дм³", value=3500.0, step=100.0, key="cust_cl_in")
        with col_cu2:
            active_c_so4 = st.number_input("SO₄²⁻, мг/дм³", value=800.0, step=50.0, key="cust_so4_in")
        active_liq_feed = calculate_liquid_waste_pollutants(q_liq, active_c_cl, active_c_so4, k_conv_cl_liq, k_conv_s_liq, dataset_name="Пользовательский")
        active_liq_title = "Пользовательский набор"

    # Отдельный расчет расхода чистого 100% NaOH на соблюдение нормативов ИТС 9-2020
    liq_calc_res = calculate_liquid_installation_naoh(
        liquid_results=active_liq_feed,
        flue_gas_flow=flue_gas_flow_liq,
        k_excess=k_excess,
        eta_scrubber=eta_scrubber,
        hours_per_day=hours_per_day,
        operating_days_year=operating_days_year
    )
    
    st.markdown("---")
    st.markdown(f"#### 📊 Результаты расчета чистого 100% NaOH ({active_liq_title}, V_г = {flue_gas_flow_liq:.0f} нм³/ч):")
    
    # KPI карточки
    l_kpi1, l_kpi2, l_kpi3 = st.columns(3)
    with l_kpi1:
        st.metric(
            label="Часовой расход 100% чистого NaOH",
            value=f"{liq_calc_res['naoh_pure_hour_kg']:.2f} кг/ч",
            delta=f"Суточный: {liq_calc_res['naoh_pure_day_kg']:.1f} кг/сут ({hours_per_day:.0f} ч/сут)"
        )
    with l_kpi2:
        st.metric(
            label="ГОДОВОЙ расход 100% чистого NaOH",
            value=f"{liq_calc_res['naoh_pure_year_t']:.2f} т/год",
            delta=f"{liq_calc_res['naoh_pure_year_t']*1000:.0f} кг/год ({operating_days_year:.0f} дн/год)"
        )
    with l_kpi3:
        st.metric(
            label="Удельный расход 100% NaOH",
            value=f"{liq_calc_res['spec_naoh_pure_g_per_kg']:.2f} г/кг",
            delta=f"{liq_calc_res['spec_naoh_pure_per_m3_kg']:.2f} кг NaOH / м³ стоков"
        )

    # Блок удельного расхода
    st.markdown(f"""
    <div class="badge-spec">
    <b>🎯 Удельный расход чистого реагента (Жидкие стоки):</b> 
    <b>{liq_calc_res['spec_naoh_pure_g_per_kg']:.2f} г 100% NaOH на 1 кг стоков</b> 
    (или <b>{liq_calc_res['spec_naoh_pure_per_m3_kg']:.2f} кг 100% NaOH на 1 м³ стоков</b>).
    </div>
    """, unsafe_allow_html=True)

    # Карточка показателей очистки газов
    st.markdown(f"""
    <div style="background-color: #f0f7fb; border: 1px solid #b8daff; border-left: 5px solid #2980b9; padding: 12px 16px; border-radius: 6px; margin: 15px 0;">
        <div style="font-weight: bold; font-size: 15px; color: #1c5980; margin-bottom: 8px;">
            💨 Параметры очистки дымовых газов: Установка жидких отходов ({flue_gas_flow_liq:.0f} нм³/ч)
        </div>
        <div style="display: flex; gap: 30px; flex-wrap: wrap; font-size: 14px;">
            <div><b>HCl вход:</b> <code>{liq_calc_res['conc_hcl_in']:.1f} мг/нм³</code> &nbsp;➔&nbsp; <b>HCl выход:</b> <code style="color: #27ae60; font-weight: bold;">{liq_calc_res['conc_hcl_out']:.1f} мг/нм³</code> <span style="color: #555;">(ПДК ≤ 10,0, η = {liq_calc_res['eta_req_hcl']:.1%})</span></div>
            <div><b>SO₂ вход:</b> <code>{liq_calc_res['conc_so2_in']:.1f} мг/нм³</code> &nbsp;➔&nbsp; <b>SO₂ выход:</b> <code style="color: #27ae60; font-weight: bold;">{liq_calc_res['conc_so2_out']:.1f} мг/нм³</code> <span style="color: #555;">(ПДК ≤ 50,0, η = {liq_calc_res['eta_req_so2']:.1%})</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Таблица материального баланса для жидких отходов
    st.markdown("##### 📋 Баланс кислых газов и реагента (Установка жидких отходов):")
    df_liq_table = pd.DataFrame([
        {
            "Компонент": "Хлороводород (HCl)",
            "Стехиометрия": "HCl + NaOH → NaCl + H₂O",
            "Поступление элемента, кг/ч": f"{active_liq_feed['mass_cl']:.3f} (Cl)",
            "k_конв": f"{k_conv_cl_liq:.2f}",
            "В газе, кг/ч": f"{active_liq_feed['mass_hcl']:.3f}",
            "Вход, мг/нм³": f"{liq_calc_res['conc_hcl_in']:.1f}",
            "ПДК, мг/нм³": "≤ 10,0",
            "Выход, мг/нм³": f"{liq_calc_res['conc_hcl_out']:.1f}",
            "Требуемая η": f"{liq_calc_res['eta_req_hcl']:.1%}",
            "100% NaOH (теор), кг/ч": f"{liq_calc_res['naoh_hcl_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{liq_calc_res['naoh_hcl_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{liq_calc_res['naoh_hcl_fact']*hours_per_day:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{(liq_calc_res['naoh_hcl_fact']*annual_hours)/1000:.2f}"
        },
        {
            "Компонент": "Диоксид серы (SO₂)",
            "Стехиометрия": "SO₂ + 2NaOH → Na₂SO₃ + H₂O",
            "Поступление элемента, кг/ч": f"{active_liq_feed['mass_s']:.3f} (S)",
            "k_конв": f"{k_conv_s_liq:.2f}",
            "В газе, кг/ч": f"{active_liq_feed['mass_so2']:.3f}",
            "Вход, мг/нм³": f"{liq_calc_res['conc_so2_in']:.1f}",
            "ПДК, мг/нм³": "≤ 50,0",
            "Выход, мг/нм³": f"{liq_calc_res['conc_so2_out']:.1f}",
            "Требуемая η": f"{liq_calc_res['eta_req_so2']:.1%}",
            "100% NaOH (теор), кг/ч": f"{liq_calc_res['naoh_so2_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{liq_calc_res['naoh_so2_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{liq_calc_res['naoh_so2_fact']*hours_per_day:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{(liq_calc_res['naoh_so2_fact']*annual_hours)/1000:.2f}"
        },
        {
            "Компонент": "ИТОГО (Жидкие отходы)",
            "Стехиометрия": "Суммарный расход",
            "Поступление элемента, кг/ч": f"{(active_liq_feed['mass_cl'] + active_liq_feed['mass_s']):.3f}",
            "k_конв": "-",
            "В газе, кг/ч": f"{(active_liq_feed['mass_hcl'] + active_liq_feed['mass_so2']):.3f}",
            "Вход, мг/нм³": "-",
            "ПДК, мг/нм³": "-",
            "Выход, мг/нм³": "-",
            "Требуемая η": "-",
            "100% NaOH (теор), кг/ч": f"{liq_calc_res['naoh_total_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{liq_calc_res['naoh_total_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{liq_calc_res['naoh_pure_day_kg']:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{liq_calc_res['naoh_pure_year_t']:.2f}"
        }
    ])
    st.dataframe(df_liq_table, use_container_width=True, hide_index=True)



# ------------------------------------------------------------------------------
# ВКЛАДКА 2: УСТАНОВКА УТИЛИЗАЦИИ ТБО
# ------------------------------------------------------------------------------
with tab2:
    st.subheader("🗑️ Установка №2: Утилизация ТБО (170 кг/ч)")
    st.markdown("Задание содержания активных кислотообразующих элементов (**Cl и S**) в ТБО:")
    
    col_tm, col_tf = st.columns([2, 1])
    with col_tf:
        m_tbo = st.number_input("Расход ТБО (M_tbo), кг/ч", min_value=10.0, max_value=2000.0, value=170.0, step=10.0, key="tbo_m_input")
        calorific_value = st.number_input("Теплота сгорания, ккал/кг", min_value=1000, max_value=8000, value=2500, step=100, key="tbo_q_input")
    with col_tm:
        tbo_method = st.radio(
            "Способ задания состава ТБО:",
            options=["morphology", "direct"],
            format_func=lambda x: {
                "morphology": "📋 По группам отходов (морфологический состав)",
                "direct": "🎯 Прямой ввод содержания Cl % и S % в смеси"
            }[x],
            key="tbo_method_rad"
        )
        
    st.markdown("---")
    
    if tbo_method == "direct":
        c_d1, c_d2 = st.columns(2)
        with c_d1:
            d_cl = st.number_input("Cl в ТБО, % масс.", value=0.35, step=0.05, key="d_cl_tbo")
        with c_d2:
            d_s = st.number_input("S в ТБО, % масс.", value=0.30, step=0.05, key="d_s_tbo")
            
        tbo_feed_res = calculate_tbo_pollutants(m_tbo, custom_elements={"cl": d_cl, "s": d_s}, k_conv_cl=k_conv_cl_tbo, k_conv_s=k_conv_s_tbo)
    else:
        if 'last_preset' not in st.session_state:
            st.session_state['last_preset'] = "Усредненный смешанный состав (рекомендуемый)"
            for g_name, g_info in TBO_WASTE_GROUPS.items():
                st.session_state[f"s_m_{g_name}"] = float(TBO_PRESETS["Усредненный смешанный состав (рекомендуемый)"].get(g_name, 0.0))

        col_pres, col_btn_reset = st.columns([2, 1])
        with col_pres:
            sel_pres = st.selectbox(
                "Пресет морфологии ТБО:",
                list(TBO_PRESETS.keys()) + ["Пользовательский"],
                key="sel_tbo_preset"
            )
        
        if sel_pres != st.session_state.get('last_preset') and sel_pres in TBO_PRESETS:
            st.session_state['last_preset'] = sel_pres
            for g_name, val in TBO_PRESETS[sel_pres].items():
                st.session_state[f"s_m_{g_name}"] = float(val)
                
        user_morph = {}
        cols_mor = st.columns(2)
        m_idx = 0
        for g_name, g_info in TBO_WASTE_GROUPS.items():
            c_target = cols_mor[m_idx % 2]
            v_cur = st.session_state.get(f"s_m_{g_name}", 0.0)
            val = c_target.slider(
                f"{g_name} (Cl: {g_info['cl']}%, S: {g_info['s']}%)",
                0.0, 100.0, float(v_cur), 1.0,
                key=f"s_m_{g_name}"
            )
            user_morph[g_name] = val
            m_idx += 1
            
        tot_m = sum(user_morph.values())
        if abs(tot_m - 100.0) > 0.01:
            st.error(f"⚠️ Сумма долей ТБО: **{tot_m:.1f}%** (должна быть 100%).")
            if st.button("Нормализовать доли ТБО до 100%"):
                for k, v in user_morph.items():
                    st.session_state[f"s_m_{k}"] = round(v / tot_m * 100.0, 1)
                st.rerun()
        else:
            st.success(f"✅ Сумма долей ТБО: **{tot_m:.1f}%** (среднее Cl = {(sum(v*TBO_WASTE_GROUPS[k]['cl'] for k,v in user_morph.items())/100):.3f}%, S = {(sum(v*TBO_WASTE_GROUPS[k]['s'] for k,v in user_morph.items())/100):.3f}%)")
            
        tbo_feed_res = calculate_tbo_pollutants(m_tbo, morphology_dict=user_morph, k_conv_cl=k_conv_cl_tbo, k_conv_s=k_conv_s_tbo)


    # Отдельный расчет расхода чистого 100% NaOH для Установки ТБО на соблюдение нормативов ИТС 9-2020
    tbo_calc_res = calculate_tbo_installation_naoh(
        tbo_results=tbo_feed_res,
        flue_gas_flow=flue_gas_flow_tbo,
        k_excess=k_excess,
        eta_scrubber=eta_scrubber,
        hours_per_day=hours_per_day,
        operating_days_year=operating_days_year
    )
    
    st.markdown("---")
    st.markdown(f"#### 📊 Результаты расчета чистого 100% NaOH (ТБО 170 кг/ч, V_г = {flue_gas_flow_tbo:.0f} нм³/ч):")
    
    t_kpi1, t_kpi2, t_kpi3 = st.columns(3)
    with t_kpi1:
        st.metric(
            label="Часовой расход 100% чистого NaOH",
            value=f"{tbo_calc_res['naoh_pure_hour_kg']:.2f} кг/ч",
            delta=f"Суточный: {tbo_calc_res['naoh_pure_day_kg']:.1f} кг/сут ({hours_per_day:.0f} ч/сут)"
        )
    with t_kpi2:
        st.metric(
            label="ГОДОВОЙ расход 100% чистого NaOH",
            value=f"{tbo_calc_res['naoh_pure_year_t']:.2f} т/год",
            delta=f"{tbo_calc_res['naoh_pure_year_t']*1000:.0f} кг/год ({operating_days_year:.0f} дн/год)"
        )
    with t_kpi3:
        st.metric(
            label="Удельный расход 100% NaOH",
            value=f"{tbo_calc_res['spec_naoh_pure_g_per_kg']:.1f} г/кг",
            delta=f"{tbo_calc_res['spec_naoh_pure_kg_per_kg']:.4f} кг NaOH / кг ТБО"
        )

    # Блок удельного расхода
    st.markdown(f"""
    <div class="badge-spec">
    <b>🎯 Удельный расход чистого реагента (ТБО):</b> 
    <b>{tbo_calc_res['spec_naoh_pure_g_per_kg']:.1f} г 100% NaOH на 1 кг ТБО</b> 
    (или <b>{tbo_calc_res['spec_naoh_pure_kg_per_kg']:.4f} кг 100% NaOH на 1 кг ТБО</b>).
    </div>
    """, unsafe_allow_html=True)

    # Карточка показателей очистки газов
    st.markdown(f"""
    <div style="background-color: #fef9f3; border: 1px solid #ffe8cc; border-left: 5px solid #e67e22; padding: 12px 16px; border-radius: 6px; margin: 15px 0;">
        <div style="font-weight: bold; font-size: 15px; color: #a0520d; margin-bottom: 8px;">
            💨 Параметры очистки дымовых газов: Установка ТБО ({flue_gas_flow_tbo:.0f} нм³/ч)
        </div>
        <div style="display: flex; gap: 30px; flex-wrap: wrap; font-size: 14px;">
            <div><b>HCl вход:</b> <code>{tbo_calc_res['conc_hcl_in']:.1f} мг/нм³</code> &nbsp;➔&nbsp; <b>HCl выход:</b> <code style="color: #27ae60; font-weight: bold;">{tbo_calc_res['conc_hcl_out']:.1f} мг/нм³</code> <span style="color: #555;">(ПДК ≤ 10,0, η = {tbo_calc_res['eta_req_hcl']:.1%})</span></div>
            <div><b>SO₂ вход:</b> <code>{tbo_calc_res['conc_so2_in']:.1f} мг/нм³</code> &nbsp;➔&nbsp; <b>SO₂ выход:</b> <code style="color: #27ae60; font-weight: bold;">{tbo_calc_res['conc_so2_out']:.1f} мг/нм³</code> <span style="color: #555;">(ПДК ≤ 50,0, η = {tbo_calc_res['eta_req_so2']:.1%})</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Таблица материального баланса для ТБО
    st.markdown("##### 📋 Баланс кислых газов и реагента (Установка ТБО):")
    df_tbo_table = pd.DataFrame([
        {
            "Компонент": "Хлороводород (HCl)",
            "Стехиометрия": "HCl + NaOH → NaCl + H₂O",
            "Поступление элемента, кг/ч": f"{tbo_feed_res['mass_cl']:.4f} (Cl)",
            "k_конв": f"{k_conv_cl_tbo:.2f}",
            "В газе, кг/ч": f"{tbo_feed_res['mass_hcl']:.3f}",
            "Вход, мг/нм³": f"{tbo_calc_res['conc_hcl_in']:.1f}",
            "ПДК, мг/нм³": "≤ 10,0",
            "Выход, мг/нм³": f"{tbo_calc_res['conc_hcl_out']:.1f}",
            "Требуемая η": f"{tbo_calc_res['eta_req_hcl']:.1%}",
            "100% NaOH (теор), кг/ч": f"{tbo_calc_res['naoh_hcl_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{tbo_calc_res['naoh_hcl_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{tbo_calc_res['naoh_hcl_fact']*hours_per_day:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{(tbo_calc_res['naoh_hcl_fact']*annual_hours)/1000:.2f}"
        },
        {
            "Компонент": "Диоксид серы (SO₂)",
            "Стехиометрия": "SO₂ + 2NaOH → Na₂SO₃ + H₂O",
            "Поступление элемента, кг/ч": f"{tbo_feed_res['mass_s']:.4f} (S)",
            "k_конв": f"{k_conv_s_tbo:.2f}",
            "В газе, кг/ч": f"{tbo_feed_res['mass_so2']:.3f}",
            "Вход, мг/нм³": f"{tbo_calc_res['conc_so2_in']:.1f}",
            "ПДК, мг/нм³": "≤ 50,0",
            "Выход, мг/нм³": f"{tbo_calc_res['conc_so2_out']:.1f}",
            "Требуемая η": f"{tbo_calc_res['eta_req_so2']:.1%}",
            "100% NaOH (теор), кг/ч": f"{tbo_calc_res['naoh_so2_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{tbo_calc_res['naoh_so2_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{tbo_calc_res['naoh_so2_fact']*hours_per_day:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{(tbo_calc_res['naoh_so2_fact']*annual_hours)/1000:.2f}"
        },
        {
            "Компонент": "ИТОГО (ТБО 170 кг/ч)",
            "Стехиометрия": "Суммарный расход",
            "Поступление элемента, кг/ч": f"{(tbo_feed_res['mass_cl'] + tbo_feed_res['mass_s']):.4f}",
            "k_конв": "-",
            "В газе, кг/ч": f"{(tbo_feed_res['mass_hcl'] + tbo_feed_res['mass_so2']):.3f}",
            "Вход, мг/нм³": "-",
            "ПДК, мг/нм³": "-",
            "Выход, мг/нм³": "-",
            "Требуемая η": "-",
            "100% NaOH (теор), кг/ч": f"{tbo_calc_res['naoh_total_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{tbo_calc_res['naoh_total_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{tbo_calc_res['naoh_pure_day_kg']:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{tbo_calc_res['naoh_pure_year_t']:.2f}"
        }
    ])
    st.dataframe(df_tbo_table, use_container_width=True, hide_index=True)



# ------------------------------------------------------------------------------
# ВКЛАДКА 3: СВОДНАЯ ВЕДОМОСТЬ И СРАВНЕНИЕ
# ------------------------------------------------------------------------------
with tab3:
    st.subheader("📊 3. Сводная ведомость и Сравнение установок")
    
    comb_calc_res = calculate_combined_installations_naoh(liq_calc_res, tbo_calc_res)
    final_results = calculate_naoh_and_compliance(
        liquid_results=active_liq_feed,
        tbo_results=tbo_feed_res,
        flue_gas_flow_liq=flue_gas_flow_liq,
        flue_gas_flow_tbo=flue_gas_flow_tbo,
        eta_scrubber=eta_scrubber,
        k_excess=k_excess,
        eta_co2_abs=0.0,
        c_naoh_sol=100.0,
        hours_per_day=hours_per_day,
        operating_days_year=operating_days_year
    )
    
    st.session_state['liq_calc'] = liq_calc_res
    st.session_state['tbo_calc'] = tbo_calc_res
    st.session_state['comb_calc'] = comb_calc_res
    st.session_state['liquid_feed'] = active_liq_feed
    st.session_state['tbo_feed'] = tbo_feed_res
    st.session_state['final_results'] = final_results
    st.session_state['params'] = {
        'q_liq': q_liq,
        'c_cl_liq': active_c_cl,
        'c_so4_liq': active_c_so4,
        'dataset_name': active_liq_title,
        'm_tbo': m_tbo,
        'calorific_value': calorific_value,
        'hours_per_day': hours_per_day,
        'operating_days_year': operating_days_year,
        'annual_hours': annual_hours,
        'eta_scrubber': eta_scrubber,
        'k_excess': k_excess,
        'flue_gas_flow_liq': flue_gas_flow_liq,
        'flue_gas_flow_tbo': flue_gas_flow_tbo,
        'flue_gas_flow': flue_gas_flow,
        'k_conv_cl_liq': k_conv_cl_liq,

        'k_conv_s_liq': k_conv_s_liq,
        'k_conv_cl_tbo': k_conv_cl_tbo,
        'k_conv_s_tbo': k_conv_s_tbo,
        'final_results': final_results
    }

    
    # === БЛОК COMPLIANCE ПО ИТС 9-2020 ДЛЯ ОБЕИХ УСТАНОВОК ===
    st.subheader("🛡️ Нормативы выбросов ИТС 9-2020 и показатели очистки газов")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.markdown(f"""
        <div style="background-color: #f0f7fb; border: 1px solid #b8daff; border-left: 5px solid #2980b9; padding: 14px 18px; border-radius: 6px; margin-bottom: 15px;">
            <div style="font-size: 16px; font-weight: bold; color: #1c5980; margin-bottom: 10px;">
                💧 Установка №1: Жидкие ({flue_gas_flow_liq:.0f} нм³/ч)
            </div>
            <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                <tr style="border-bottom: 1px solid #d0e1f0;">
                    <td style="padding: 6px 0;"><b>HCl вход:</b> <code>{liq_calc_res['conc_hcl_in']:.1f} мг/нм³</code></td>
                    <td style="padding: 6px 0;"><b>HCl выход:</b> <code style="color: #27ae60; font-weight: bold;">{liq_calc_res['conc_hcl_out']:.1f} мг/нм³</code> <span style="color: #666; font-size: 12px;">(ПДК ≤ 10,0)</span></td>
                </tr>
                <tr>
                    <td style="padding: 6px 0;"><b>SO₂ вход:</b> <code>{liq_calc_res['conc_so2_in']:.1f} мг/нм³</code></td>
                    <td style="padding: 6px 0;"><b>SO₂ выход:</b> <code style="color: #27ae60; font-weight: bold;">{liq_calc_res['conc_so2_out']:.1f} мг/нм³</code> <span style="color: #666; font-size: 12px;">(ПДК ≤ 50,0)</span></td>
                </tr>
            </table>
            <div style="margin-top: 8px; font-size: 12px; color: #555;">
                Требуемая глубина очистки под ПДК: HCl = <b>{liq_calc_res['eta_req_hcl']:.1%}</b>, SO₂ = <b>{liq_calc_res['eta_req_so2']:.1%}</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_u2:
        st.markdown(f"""
        <div style="background-color: #fef9f3; border: 1px solid #ffe8cc; border-left: 5px solid #e67e22; padding: 14px 18px; border-radius: 6px; margin-bottom: 15px;">
            <div style="font-size: 16px; font-weight: bold; color: #a0520d; margin-bottom: 10px;">
                🗑️ Установка №2: ТБО ({flue_gas_flow_tbo:.0f} нм³/ч)
            </div>
            <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                <tr style="border-bottom: 1px solid #fae2cc;">
                    <td style="padding: 6px 0;"><b>HCl вход:</b> <code>{tbo_calc_res['conc_hcl_in']:.1f} мг/нм³</code></td>
                    <td style="padding: 6px 0;"><b>HCl выход:</b> <code style="color: #27ae60; font-weight: bold;">{tbo_calc_res['conc_hcl_out']:.1f} мг/нм³</code> <span style="color: #666; font-size: 12px;">(ПДК ≤ 10,0)</span></td>
                </tr>
                <tr>
                    <td style="padding: 6px 0;"><b>SO₂ вход:</b> <code>{tbo_calc_res['conc_so2_in']:.1f} мг/нм³</code></td>
                    <td style="padding: 6px 0;"><b>SO₂ выход:</b> <code style="color: #27ae60; font-weight: bold;">{tbo_calc_res['conc_so2_out']:.1f} мг/нм³</code> <span style="color: #666; font-size: 12px;">(ПДК ≤ 50,0)</span></td>
                </tr>
            </table>
            <div style="margin-top: 8px; font-size: 12px; color: #555;">
                Требуемая глубина очистки под ПДК: HCl = <b>{tbo_calc_res['eta_req_hcl']:.1%}</b>, SO₂ = <b>{tbo_calc_res['eta_req_so2']:.1%}</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

    if final_results['compliance_msg']:
        st.warning(final_results['compliance_msg'])
    else:
        st.success(f"✅ Расход NaOH точно подобран для достижения нормативов ИТС 9-2020 (HCl ≤ 10,0 мг/нм³, SO₂ ≤ 50,0 мг/нм³). Паспортная способность скрубберов ({eta_scrubber:.1%}) достаточна.")

    # Таблица показателей на выходе (Compliance ИТС 9-2020)
    st.markdown("##### 📋 Сводная таблица входных и выходных параметров газов (ИТС 9-2020):")
    df_comp_table = pd.DataFrame([
        {
            "Установка": "💧 Уст. №1 (Жидкие отходы, 1,5 м³/ч)",
            "Загрязняющее вещество": "Хлороводород (HCl)",
            "Расход газов, нм³/ч": f"{flue_gas_flow_liq:.0f}",
            "Вход, мг/нм³": f"{liq_calc_res['conc_hcl_in']:.1f}",
            "ПДК ИТС 9-2020, мг/нм³": "≤ 10,0",
            "Выход, мг/нм³": f"{liq_calc_res['conc_hcl_out']:.1f}",
            "Требуемая η": f"{liq_calc_res['eta_req_hcl']:.1%}",
            "Выброс в трубу, кг/ч": f"{((liq_calc_res['conc_hcl_out']*flue_gas_flow_liq)/1e6):.4f}",
            "Статус": "✅ НОРМА" if liq_calc_res['conc_hcl_out'] <= 10.0 else "⚠️ ПРЕВЫШЕНИЕ"
        },
        {
            "Установка": "💧 Уст. №1 (Жидкие отходы, 1,5 м³/ч)",
            "Загрязняющее вещество": "Диоксид серы (SO₂)",
            "Расход газов, нм³/ч": f"{flue_gas_flow_liq:.0f}",
            "Вход, мг/нм³": f"{liq_calc_res['conc_so2_in']:.1f}",
            "ПДК ИТС 9-2020, мг/нм³": "≤ 50,0",
            "Выход, мг/нм³": f"{liq_calc_res['conc_so2_out']:.1f}",
            "Требуемая η": f"{liq_calc_res['eta_req_so2']:.1%}",
            "Выброс в трубу, кг/ч": f"{((liq_calc_res['conc_so2_out']*flue_gas_flow_liq)/1e6):.4f}",
            "Статус": "✅ НОРМА" if liq_calc_res['conc_so2_out'] <= 50.0 else "⚠️ ПРЕВЫШЕНИЕ"
        },
        {
            "Установка": "🗑️ Уст. №2 (ТБО, 170 кг/ч)",
            "Загрязняющее вещество": "Хлороводород (HCl)",
            "Расход газов, нм³/ч": f"{flue_gas_flow_tbo:.0f}",
            "Вход, мг/нм³": f"{tbo_calc_res['conc_hcl_in']:.1f}",
            "ПДК ИТС 9-2020, мг/нм³": "≤ 10,0",
            "Выход, мг/нм³": f"{tbo_calc_res['conc_hcl_out']:.1f}",
            "Требуемая η": f"{tbo_calc_res['eta_req_hcl']:.1%}",
            "Выброс в трубу, кг/ч": f"{((tbo_calc_res['conc_hcl_out']*flue_gas_flow_tbo)/1e6):.4f}",
            "Статус": "✅ НОРМА" if tbo_calc_res['conc_hcl_out'] <= 10.0 else "⚠️ ПРЕВЫШЕНИЕ"
        },
        {
            "Установка": "🗑️ Уст. №2 (ТБО, 170 кг/ч)",
            "Загрязняющее вещество": "Диоксид серы (SO₂)",
            "Расход газов, нм³/ч": f"{flue_gas_flow_tbo:.0f}",
            "Вход, мг/нм³": f"{tbo_calc_res['conc_so2_in']:.1f}",
            "ПДК ИТС 9-2020, мг/нм³": "≤ 50,0",
            "Выход, мг/нм³": f"{tbo_calc_res['conc_so2_out']:.1f}",
            "Требуемая η": f"{tbo_calc_res['eta_req_so2']:.1%}",
            "Выброс в трубу, кг/ч": f"{((tbo_calc_res['conc_so2_out']*flue_gas_flow_tbo)/1e6):.4f}",
            "Статус": "✅ НОРМА" if tbo_calc_res['conc_so2_out'] <= 50.0 else "⚠️ ПРЕВЫШЕНИЕ"
        }

    ])
    st.dataframe(df_comp_table, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown(f"**График работы:** {hours_per_day:.0f} ч/сутки • {operating_days_year:.0f} рабочих дней в году • Общий фонд: **{annual_hours:.0f} ч/год**")

    # Сводная таблица
    st.markdown("##### 📊 Сводная ведомость потребности в чистом 100% NaOH и показателей выбросов:")
    mass_hcl_out_l = (liq_calc_res['conc_hcl_out'] * flue_gas_flow_liq) / 1e6
    mass_so2_out_l = (liq_calc_res['conc_so2_out'] * flue_gas_flow_liq) / 1e6
    mass_hcl_out_t = (tbo_calc_res['conc_hcl_out'] * flue_gas_flow_tbo) / 1e6
    mass_so2_out_t = (tbo_calc_res['conc_so2_out'] * flue_gas_flow_tbo) / 1e6
    conc_hcl_out_comb = ((mass_hcl_out_l + mass_hcl_out_t) * 1e6) / flue_gas_flow if flue_gas_flow > 0 else 0.0
    conc_so2_out_comb = ((mass_so2_out_l + mass_so2_out_t) * 1e6) / flue_gas_flow if flue_gas_flow > 0 else 0.0

    df_compare_installations = pd.DataFrame([
        {
            "Показатель": "Расход дымовых газов (V_г)",
            "Ед. изм.": "нм³/ч",
            "1) Жидкие отходы (1,5 м³/ч)": f"{flue_gas_flow_liq:.0f}",
            "2) ТБО (170 кг/ч)": f"{flue_gas_flow_tbo:.0f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{flue_gas_flow:.0f}"
        },
        {
            "Показатель": "HCl: входная концентрация",
            "Ед. изм.": "мг/нм³",
            "1) Жидкие отходы (1,5 м³/ч)": f"{liq_calc_res['conc_hcl_in']:.1f}",
            "2) ТБО (170 кг/ч)": f"{tbo_calc_res['conc_hcl_in']:.1f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{final_results['conc_hcl_in']:.1f}"
        },
        {
            "Показатель": "HCl: выходная концентрация (ПДК ≤ 10,0)",
            "Ед. изм.": "мг/нм³",
            "1) Жидкие отходы (1,5 м³/ч)": f"{liq_calc_res['conc_hcl_out']:.1f}",
            "2) ТБО (170 кг/ч)": f"{tbo_calc_res['conc_hcl_out']:.1f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{conc_hcl_out_comb:.1f}"
        },
        {
            "Показатель": "SO₂: входная концентрация",
            "Ед. изм.": "мг/нм³",
            "1) Жидкие отходы (1,5 м³/ч)": f"{liq_calc_res['conc_so2_in']:.1f}",
            "2) ТБО (170 кг/ч)": f"{tbo_calc_res['conc_so2_in']:.1f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{final_results['conc_so2_in']:.1f}"
        },
        {
            "Показатель": "SO₂: выходная концентрация (ПДК ≤ 50,0)",
            "Ед. изм.": "мг/нм³",
            "1) Жидкие отходы (1,5 м³/ч)": f"{liq_calc_res['conc_so2_out']:.1f}",
            "2) ТБО (170 кг/ч)": f"{tbo_calc_res['conc_so2_out']:.1f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{conc_so2_out_comb:.1f}"
        },
        {
            "Показатель": "Часовой расход чистого 100% NaOH",
            "Ед. изм.": "кг/ч",
            "1) Жидкие отходы (1,5 м³/ч)": f"{liq_calc_res['naoh_pure_hour_kg']:.2f}",
            "2) ТБО (170 кг/ч)": f"{tbo_calc_res['naoh_pure_hour_kg']:.2f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{comb_calc_res['naoh_pure_hour_kg']:.2f}"
        },
        {
            "Показатель": f"Суточный расход 100% NaOH ({hours_per_day:.0f} ч/сут)",
            "Ед. изм.": "кг/сут",
            "1) Жидкие отходы (1,5 м³/ч)": f"{liq_calc_res['naoh_pure_day_kg']:.1f}",
            "2) ТБО (170 кг/ч)": f"{tbo_calc_res['naoh_pure_day_kg']:.1f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{comb_calc_res['naoh_pure_day_kg']:.1f}"
        },
        {
            "Показатель": f"ГОДОВОЙ РАСХОД ЧИСТОГО 100% NaOH ({operating_days_year:.0f} дн/год)",
            "Ед. изм.": "т/год",
            "1) Жидкие отходы (1,5 м³/ч)": f"{liq_calc_res['naoh_pure_year_t']:.2f}",
            "2) ТБО (170 кг/ч)": f"{tbo_calc_res['naoh_pure_year_t']:.2f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{comb_calc_res['naoh_pure_year_t']:.2f}"
        },
        {
            "Показатель": "УДЕЛЬНЫЙ РАСХОД 100% NaOH НА 1 КГ ОТХОДА",
            "Ед. изм.": "г NaOH / кг отхода",
            "1) Жидкие отходы (1,5 м³/ч)": f"{liq_calc_res['spec_naoh_pure_g_per_kg']:.2f}",
            "2) ТБО (170 кг/ч)": f"{tbo_calc_res['spec_naoh_pure_g_per_kg']:.1f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{comb_calc_res['spec_naoh_pure_g_per_kg']:.1f}"
        },
        {
            "Показатель": "Остаточный выброс HCl в атмосферу",
            "Ед. изм.": "кг/ч",
            "1) Жидкие отходы (1,5 м³/ч)": f"{mass_hcl_out_l:.4f}",
            "2) ТБО (170 кг/ч)": f"{mass_hcl_out_t:.4f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{(mass_hcl_out_l + mass_hcl_out_t):.4f}"
        },
        {
            "Показатель": "Остаточный выброс SO₂ в атмосферу",
            "Ед. изм.": "кг/ч",
            "1) Жидкие отходы (1,5 м³/ч)": f"{mass_so2_out_l:.4f}",
            "2) ТБО (170 кг/ч)": f"{mass_so2_out_t:.4f}",
            "СУММАРНО ПО КОМПЛЕКСУ": f"{(mass_so2_out_l + mass_so2_out_t):.4f}"
        }
    ])
    st.dataframe(df_compare_installations, use_container_width=True, hide_index=True)


    
    st.markdown("---")
    
    # Графики сравнения
    cg1, cg2 = st.columns(2)
    with cg1:
        fig_ann = px.bar(
            x=["Жидкие отходы", "ТБО", "СУММАРНО"],
            y=[liq_calc_res['naoh_pure_year_t'], tbo_calc_res['naoh_pure_year_t'], comb_calc_res['naoh_pure_year_t']],
            color=["Жидкие отходы", "ТБО", "СУММАРНО"],
            color_discrete_sequence=['#2980B9', '#E67E22', '#27AE60'],
            title=f"Годовой расход чистого 100% NaOH (т/год при {operating_days_year:.0f} дн/год)",
            labels={"x": "Установка", "y": "т/год чистого NaOH"}
        )
        fig_ann.update_layout(showlegend=False, height=320, margin=dict(t=40, b=10, l=10, r=10))
        st.plotly_chart(fig_ann, use_container_width=True)
        
    with cg2:
        fig_spec = px.bar(
            x=["Жидкие отходы", "ТБО", "СУММАРНО"],
            y=[liq_calc_res['spec_naoh_pure_g_per_kg'], tbo_calc_res['spec_naoh_pure_g_per_kg'], comb_calc_res['spec_naoh_pure_g_per_kg']],
            color=["Жидкие отходы", "ТБО", "СУММАРНО"],
            color_discrete_sequence=['#2980B9', '#E67E22', '#27AE60'],
            title="Удельный расход 100% NaOH на 1 кг отхода (г/кг)",
            labels={"x": "Установка", "y": "г NaOH / кг отхода"}
        )
        fig_spec.update_layout(showlegend=False, height=320, margin=dict(t=40, b=10, l=10, r=10))
        st.plotly_chart(fig_spec, use_container_width=True)

# ------------------------------------------------------------------------------
# ВКЛАДКА 4: ПОЯСНИТЕЛЬНАЯ ЗАПИСКА И ЭКСПОРТ
# ------------------------------------------------------------------------------
with tab4:
    st.subheader("📄 4. Пояснительная записка и Экспорт документации")
    st.markdown("Формирование раздельных отчетов по установкам и сводной ведомости с исходными данными, составом отходов и формулами.")
    
    if 'liq_calc' in st.session_state:
        e1, e2, e3 = st.columns(3)
        with e1:
            st.markdown("##### 📝 Microsoft Word (.docx)")
            st.caption("Полнотекстовая ПЗ с исходными данными, составом отходов и формулами.")
            docx_data = generate_word_report(
                liq_calc=st.session_state['liq_calc'],
                tbo_calc=st.session_state['tbo_calc'],
                comb_calc=st.session_state['comb_calc'],
                liquid_feed=st.session_state['liquid_feed'],
                tbo_feed=st.session_state['tbo_feed'],
                params=st.session_state['params']
            )
            st.download_button(
                label="📥 Скачать отчет Word (.docx)",
                data=docx_data,
                file_name=f"ПЗ_Расход_NaOH_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
                
        with e2:
            st.markdown("##### 📑 Документ PDF (.pdf)")
            st.caption("Инженерная справка с формулами, исходными данными и сводкой.")
            pdf_data = generate_pdf_report(
                liq_calc=st.session_state['liq_calc'],
                tbo_calc=st.session_state['tbo_calc'],
                comb_calc=st.session_state['comb_calc'],
                liquid_feed=st.session_state['liquid_feed'],
                tbo_feed=st.session_state['tbo_feed'],
                params=st.session_state['params']
            )
            st.download_button(
                label="📥 Скачать отчет PDF (.pdf)",
                data=pdf_data,
                file_name=f"ПЗ_Расход_NaOH_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                type="secondary"
            )
                
        with e3:
            st.markdown("##### 📊 Таблица Excel (.xlsx)")
            st.caption("Книга Excel: Сводка, Исходные данные и формулы, Балансы, Морфология.")
            xlsx_data = generate_excel_report(
                liq_calc=st.session_state['liq_calc'],
                tbo_calc=st.session_state['tbo_calc'],
                comb_calc=st.session_state['comb_calc'],
                liquid_feed=st.session_state['liquid_feed'],
                tbo_feed=st.session_state['tbo_feed'],
                params=st.session_state['params']
            )
            st.download_button(
                label="📥 Скачать таблицу Excel (.xlsx)",
                data=xlsx_data,
                file_name=f"Материальный_баланс_NaOH_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
                
        st.markdown("---")
        
        # Интерактивный блок формул и методики
        with st.expander("📐 Методика и формулы расчета (включены в отчет)", expanded=True):
            st.markdown(r"""
            #### 1. Образование кислых газов:
            * **Хлороводород ($\text{HCl}$)**:
              $$M_{\text{HCl}} = M_{\text{Cl}} \cdot k_{\text{конв, Cl}} \cdot \frac{\mu_{\text{HCl}}}{\mu_{\text{Cl}}} = M_{\text{Cl}} \cdot 0{,}98 \cdot \frac{36{,}461}{35{,}453} \approx M_{\text{Cl}} \cdot 1{,}0078\quad (\text{кг/ч})$$
            * **Диоксид серы ($\text{SO}_2$)**:
              $$M_{\text{SO}_2} = M_{\text{S}} \cdot k_{\text{конв, S}} \cdot \frac{\mu_{\text{SO}_2}}{\mu_{\text{S}}} = M_{\text{S}} \cdot 0{,}90 \cdot \frac{64{,}063}{32{,}065} \approx M_{\text{S}} \cdot 1{,}7981\quad (\text{кг/ч})$$
              *(для жидких отходов $M_{\text{S}} = Q_{\text{liq}} \cdot C_{\text{SO}_4} \cdot \frac{32{,}065}{96{,}061} \cdot 10^{-3}$)*

            #### 2. Стехиометрические реакции и теоретический расход $100\%$ $\text{NaOH}$:
            * **Нейтрализация $\text{HCl}$ (улавливаемая масса $M_{\text{HCl}} \cdot \eta_{\text{скр}}$)**:
              $$\text{HCl} + \text{NaOH} \rightarrow \text{NaCl} + \text{H}_2\text{O}$$
              $$M_{\text{NaOH, HCl}}^{\text{теор}} = M_{\text{HCl}} \cdot \eta_{\text{скр}} \cdot \frac{\mu_{\text{NaOH}}}{\mu_{\text{HCl}}} \approx M_{\text{HCl}} \cdot \eta_{\text{скр}} \cdot 1{,}0970\quad (\text{кг/ч})$$
            * **Нейтрализация $\text{SO}_2$ (улавливаемая масса $M_{\text{SO}_2} \cdot \eta_{\text{скр}}$)**:
              $$\text{SO}_2 + 2\text{NaOH} \rightarrow \text{Na}_2\text{SO}_3 + \text{H}_2\text{O}$$
              $$M_{\text{NaOH, SO}_2}^{\text{теор}} = M_{\text{SO}_2} \cdot \eta_{\text{скр}} \cdot \frac{2 \cdot \mu_{\text{NaOH}}}{\mu_{\text{SO}_2}} \approx M_{\text{SO}_2} \cdot \eta_{\text{скр}} \cdot 1{,}2487\quad (\text{кг/ч})$$
            * **Суммарный теоретический расход**:
              $$M_{\text{NaOH}}^{\text{теор}} = M_{\text{NaOH, HCl}}^{\text{теор}} + M_{\text{NaOH, SO}_2}^{\text{теор}}\quad (\text{кг/ч})$$

            #### 3. Фактический часовой, суточный и годовой расход чистого $100\%$ $\text{NaOH}$:
            * **Физическое обоснование**:
              *Эффективность скруббера ($\eta_{\text{скр}}$) определяет долю кислых газов, которая фактически поглощается орошающим раствором. Расход $\text{NaOH}$ рассчитывается строго исходя из массы уловленных загрязнителей. Коэффициент избытка ($k_{\text{изб}} = 1{,}15$) обеспечивает технологический запас и стабильное поддержание $\text{pH}$ на уровне 7,5–8,5.*
            * **Часовой расход с учетом КПД скруббера $\eta_{\text{скр}}$ и избытка $k_{\text{изб}}$**:
              $$M_{\text{NaOH, факт}}^{\text{час}} = M_{\text{NaOH}}^{\text{теор}} \cdot k_{\text{изб}} = \left(M_{\text{HCl}} \cdot 1{,}0970 + M_{\text{SO}_2} \cdot 1{,}2487\right) \cdot \eta_{\text{скр}} \cdot k_{\text{изб}}\quad (\text{кг/ч})$$
            * **Суточный расход по рабочей смене ($T_{\text{сут}}$)**:
              $$M_{\text{NaOH, факт}}^{\text{сут}} = M_{\text{NaOH, факт}}^{\text{час}} \cdot T_{\text{сут}}\quad (\text{кг/сут})$$
            * **Годовой расход с учетом рабочих дней ($D_{\text{год}}$)**:
              $$M_{\text{NaOH, факт}}^{\text{год}} = \frac{M_{\text{NaOH, факт}}^{\text{час}} \cdot (T_{\text{сут}} \cdot D_{\text{год}})}{1000} = \frac{M_{\text{NaOH, факт}}^{\text{час}} \cdot T_{\text{год}}}{1000}\quad (\text{т/год})$$


            #### 4. Удельный расход чистого $100\%$ реагента на 1 кг отходов:
            $$q_{\text{NaOH}} = \frac{M_{\text{NaOH, факт}}^{\text{час}}}{M_{\text{отходов}}^{\text{час}}} \cdot 1000\quad (\text{г 100\% NaOH / кг отходов})$$
            $$q_{\text{NaOH, объем}} = \frac{M_{\text{NaOH, факт}}^{\text{час}}}{Q_{\text{liq}}}\quad (\text{кг 100\% NaOH / м}^3\text{ стоков})$$
            """)
            
        with st.expander("👁️ Предпросмотр текста заключения", expanded=False):
            lc = st.session_state['liq_calc']
            tc = st.session_state['tbo_calc']
            cc = st.session_state['comb_calc']
            pm = st.session_state['params']
            fin = st.session_state.get('final_results', {})
            st.markdown(f"""
            ### ЗАКЛЮЧЕНИЕ
            
            **1. Установка утилизации жидких отходов (1,5 м³/ч, {pm['dataset_name']}):**
            - **Удельный расход: {lc['spec_naoh_pure_g_per_kg']:.2f} г чистого 100% NaOH / кг стоков** ({lc['spec_naoh_pure_per_m3_kg']:.2f} кг NaOH / м³ стоков).
            - Часовой расход чистого 100% NaOH: **{lc['naoh_pure_hour_kg']:.2f} кг/ч**.
            - Суточный расход при {pm['hours_per_day']:.0f} ч/сут: **{lc['naoh_pure_day_kg']:.1f} кг/сут**.
            - **Годовой расход чистого 100% реагента (при {pm['operating_days_year']:.0f} дн/год): {lc['naoh_pure_year_t']:.2f} т/год**.
            
            **2. Установка утилизации ТБО (170 кг/ч):**
            - **Удельный расход: {tc['spec_naoh_pure_g_per_kg']:.1f} г чистого 100% NaOH / кг ТБО** ({tc['spec_naoh_pure_kg_per_kg']:.4f} кг NaOH / кг ТБО).
            - Часовой расход чистого 100% NaOH: **{tc['naoh_pure_hour_kg']:.2f} кг/ч**.
            - Суточный расход при {pm['hours_per_day']:.0f} ч/сут: **{tc['naoh_pure_day_kg']:.1f} кг/сут**.
            - **Годовой расход чистого 100% реагента (при {pm['operating_days_year']:.0f} дн/год): {tc['naoh_pure_year_t']:.2f} т/год**.
            
            **3. Суммарная годовая потребность комплекса:**
            - **Чистый 100% NaOH: {cc['naoh_pure_year_t']:.2f} т/год** ({cc['naoh_pure_day_kg']:.1f} кг/сут).
            - **Средневзвешенный удельный расход по комплексу: {cc['spec_naoh_pure_g_per_kg']:.1f} г 100% NaOH / кг суммарных отходов**.
            
            **4. Соответствие нормативам ИТС 9-2020 (Приложение В):**
            - Расчетная концентрация HCl на выходе: **{fin.get('conc_hcl_out', 0.0):.2f} мг/нм³** (норматив ≤ 10 мг/нм³).
            - Расчетная концентрация SO₂ на выходе: **{fin.get('conc_so2_out', 0.0):.2f} мг/нм³** (норматив ≤ 50 мг/нм³).
            - Статус соответствия: **{fin.get('compliance_status', '✅ НОРМА')}**.
            """)

# Футер
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 0.9em;'>"
    "Расчет выполнен с учетом коэффициентов связывания гетероатомов в золе при Т = 1100 °C "
    "и предельных значений выбросов ИТС 9-2020 (Приложение В)."
    "</div>",
    unsafe_allow_html=True
)

