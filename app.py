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
st.sidebar.header("⚙️ Технологические параметры")

st.sidebar.subheader("Дымовые газы")
flue_gas_flow = st.sidebar.number_input(
    "Расход дымовых газов, нм³/ч",
    min_value=100.0, max_value=50000.0, value=3000.0, step=100.0,
    help="Определяет объем газов, в котором рассчитываются концентрации загрязнителей. Расход NaOH рассчитывается напрямую для достижения нормативов ИТС 9-2020 (HCl ≤ 10 мг/нм³, SO₂ ≤ 50 мг/нм³)"
)

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
    k_conv_c = st.slider("Доля C → CO₂", 0.0, 1.0, 0.98, 0.01, help="Степень полного окисления углерода в CO₂")

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
        flue_gas_flow=flue_gas_flow,
        k_excess=k_excess,
        eta_scrubber=eta_scrubber,
        hours_per_day=hours_per_day,
        operating_days_year=operating_days_year,
        total_feed_mass=total_feed_mass_est
    )
    
    st.markdown("---")
    st.markdown(f"#### 📊 Результаты расчета чистого 100% NaOH ({active_liq_title}):")
    
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

    # Таблица материального баланса для жидких отходов
    st.markdown("##### 📋 Баланс кислых газов и реагента (Установка жидких отходов):")
    df_liq_table = pd.DataFrame([
        {
            "Компонент": "Хлороводород (HCl)",
            "Стехиометрия": "HCl + NaOH → NaCl + H₂O",
            "Поступление Cl, кг/ч": f"{active_liq_feed['mass_cl']:.3f}",
            "Масса HCl в газе, кг/ч": f"{active_liq_feed['mass_hcl']:.3f}",
            "100% NaOH (теор), кг/ч": f"{liq_calc_res['naoh_hcl_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{liq_calc_res['naoh_hcl_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{liq_calc_res['naoh_hcl_fact']*hours_per_day:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{(liq_calc_res['naoh_hcl_fact']*annual_hours)/1000:.2f}",
            "Удельный, г/кг": f"{(liq_calc_res['naoh_hcl_fact']/active_liq_feed['feed_mass_kg_h']*1000):.2f}"
        },
        {
            "Компонент": "Диоксид серы (SO₂)",
            "Стехиометрия": "SO₂ + 2NaOH → Na₂SO₃ + H₂O",
            "Поступление S, кг/ч": f"{active_liq_feed['mass_s']:.3f}",
            "Масса SO₂ в газе, кг/ч": f"{active_liq_feed['mass_so2']:.3f}",
            "100% NaOH (теор), кг/ч": f"{liq_calc_res['naoh_so2_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{liq_calc_res['naoh_so2_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{liq_calc_res['naoh_so2_fact']*hours_per_day:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{(liq_calc_res['naoh_so2_fact']*annual_hours)/1000:.2f}",
            "Удельный, г/кг": f"{(liq_calc_res['naoh_so2_fact']/active_liq_feed['feed_mass_kg_h']*1000):.2f}"
        },
        {
            "Компонент": "ИТОГО (Жидкие отходы)",
            "Стехиометрия": "Суммарный расход",
            "Поступление Cl, кг/ч": "-",
            "Масса HCl в газе, кг/ч": f"{(active_liq_feed['mass_hcl'] + active_liq_feed['mass_so2']):.3f}",
            "100% NaOH (теор), кг/ч": f"{liq_calc_res['naoh_total_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{liq_calc_res['naoh_total_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{liq_calc_res['naoh_pure_day_kg']:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{liq_calc_res['naoh_pure_year_t']:.2f}",
            "Удельный, г/кг": f"{liq_calc_res['spec_naoh_pure_g_per_kg']:.2f}"
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
            
        tbo_feed_res = calculate_tbo_pollutants(m_tbo, custom_elements={"cl": d_cl, "s": d_s}, k_conv_cl=k_conv_cl_tbo, k_conv_s=k_conv_s_tbo, k_conv_c=k_conv_c)
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
            
        tbo_feed_res = calculate_tbo_pollutants(m_tbo, morphology_dict=user_morph, k_conv_cl=k_conv_cl_tbo, k_conv_s=k_conv_s_tbo, k_conv_c=k_conv_c)


    # Отдельный расчет расхода чистого 100% NaOH для Установки ТБО на соблюдение нормативов ИТС 9-2020
    tbo_calc_res = calculate_tbo_installation_naoh(
        tbo_results=tbo_feed_res,
        flue_gas_flow=flue_gas_flow,
        k_excess=k_excess,
        eta_scrubber=eta_scrubber,
        hours_per_day=hours_per_day,
        operating_days_year=operating_days_year,
        total_feed_mass=total_feed_mass_est
    )
    
    st.markdown("---")
    st.markdown("#### 📊 Результаты расчета чистого 100% NaOH (ТБО 170 кг/ч):")
    
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

    # Таблица материального баланса для ТБО
    st.markdown("##### 📋 Баланс кислых газов и реагента (Установка ТБО):")
    df_tbo_table = pd.DataFrame([
        {
            "Компонент": "Хлороводород (HCl)",
            "Стехиометрия": "HCl + NaOH → NaCl + H₂O",
            "Поступление элемента, кг/ч": f"{tbo_feed_res['mass_cl']:.4f} (Cl)",
            "Масса в газе, кг/ч": f"{tbo_feed_res['mass_hcl']:.3f}",
            "100% NaOH (теор), кг/ч": f"{tbo_calc_res['naoh_hcl_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{tbo_calc_res['naoh_hcl_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{tbo_calc_res['naoh_hcl_fact']*hours_per_day:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{(tbo_calc_res['naoh_hcl_fact']*annual_hours)/1000:.2f}",
            "Удельный, г/кг": f"{(tbo_calc_res['naoh_hcl_fact']/m_tbo*1000):.2f}"
        },
        {
            "Компонент": "Диоксид серы (SO₂)",
            "Стехиометрия": "SO₂ + 2NaOH → Na₂SO₃ + H₂O",
            "Поступление элемента, кг/ч": f"{tbo_feed_res['mass_s']:.4f} (S)",
            "Масса в газе, кг/ч": f"{tbo_feed_res['mass_so2']:.3f}",
            "100% NaOH (теор), кг/ч": f"{tbo_calc_res['naoh_so2_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{tbo_calc_res['naoh_so2_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{tbo_calc_res['naoh_so2_fact']*hours_per_day:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{(tbo_calc_res['naoh_so2_fact']*annual_hours)/1000:.2f}",
            "Удельный, г/кг": f"{(tbo_calc_res['naoh_so2_fact']/m_tbo*1000):.2f}"
        },
        {
            "Компонент": "ИТОГО (ТБО 170 кг/ч)",
            "Стехиометрия": "Суммарный расход",
            "Поступление элемента, кг/ч": "-",
            "Масса в газе, кг/ч": f"{(tbo_feed_res['mass_hcl'] + tbo_feed_res['mass_so2']):.3f}",
            "100% NaOH (теор), кг/ч": f"{tbo_calc_res['naoh_total_theor']:.3f}",
            "100% NaOH (факт), кг/ч": f"{tbo_calc_res['naoh_total_fact']:.3f}",
            "Суточный (100%), кг/сут": f"{tbo_calc_res['naoh_pure_day_kg']:.1f}",
            "ГОДОВОЙ (100%), т/год": f"{tbo_calc_res['naoh_pure_year_t']:.2f}",
            "Удельный, г/кг": f"{tbo_calc_res['spec_naoh_pure_g_per_kg']:.1f}"
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
        flue_gas_flow=flue_gas_flow,
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
        'flue_gas_flow': flue_gas_flow,
        'k_conv_cl_liq': k_conv_cl_liq,
        'k_conv_s_liq': k_conv_s_liq,
        'k_conv_cl_tbo': k_conv_cl_tbo,
        'k_conv_s_tbo': k_conv_s_tbo,
        'k_conv_c': k_conv_c,
        'final_results': final_results
    }
    
    # === БЛОК COMPLIANCE ПО ИТС 9-2020 ===
    st.subheader("🛡️ Проверка compliance по ИТС 9-2020 (Приложение В)")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("HCl на входе в скруббер", f"{final_results['conc_hcl_in']:.1f} мг/нм³")
        st.caption(f"Требуемая эффективность для выхода на ≤10 мг/нм³: **{final_results['eta_req_hcl']:.1%}**")

    with col2:
        st.metric("SO₂ на входе в скруббер", f"{final_results['conc_so2_in']:.1f} мг/нм³")
        st.caption(f"Требуемая эффективность для выхода на ≤50 мг/нм³: **{final_results['eta_req_so2']:.1%}**")

    with col3:
        st.markdown(f"### {final_results['compliance_status']}")
        if final_results['compliance_msg']:
            st.warning(final_results['compliance_msg'])
        else:
            st.success("Текущая эффективность скруббера достаточна для соблюдения нормативов.")

    st.markdown("---")
    st.markdown(f"**График работы:** {hours_per_day:.0f} ч/сутки • {operating_days_year:.0f} рабочих дней в году • Общий фонд: **{annual_hours:.0f} ч/год**")

    
    # Сводная таблица
    df_compare_installations = pd.DataFrame([
        {
            "Показатель": "Часовой расход 100% чистого NaOH",
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
            * **Нейтрализация $\text{HCl}$**:
              $$\text{HCl} + \text{NaOH} \rightarrow \text{NaCl} + \text{H}_2\text{O}$$
              $$M_{\text{NaOH, HCl}}^{\text{теор}} = M_{\text{HCl}} \cdot \frac{\mu_{\text{NaOH}}}{\mu_{\text{HCl}}} = M_{\text{HCl}} \cdot \frac{39{,}997}{36{,}461} \approx M_{\text{HCl}} \cdot 1{,}0970\quad (\text{кг/ч})$$
            * **Нейтрализация $\text{SO}_2$**:
              $$\text{SO}_2 + 2\text{NaOH} \rightarrow \text{Na}_2\text{SO}_3 + \text{H}_2\text{O}$$
              $$M_{\text{NaOH, SO}_2}^{\text{теор}} = M_{\text{SO}_2} \cdot \frac{2 \cdot \mu_{\text{NaOH}}}{\mu_{\text{SO}_2}} = M_{\text{SO}_2} \cdot \frac{79{,}994}{64{,}063} \approx M_{\text{SO}_2} \cdot 1{,}2487\quad (\text{кг/ч})$$
            * **Суммарный теоретический расход**:
              $$M_{\text{NaOH}}^{\text{теор}} = M_{\text{NaOH, HCl}}^{\text{теор}} + M_{\text{NaOH, SO}_2}^{\text{теор}}\quad (\text{кг/ч})$$

            #### 3. Фактический часовой, суточный и годовой расход чистого $100\%$ $\text{NaOH}$:
            * **Часовой расход с учетом КПД скруббера $\eta_{\text{скр}}$ и избытка $k_{\text{изб}}$**:
              $$M_{\text{NaOH, факт}}^{\text{час}} = M_{\text{NaOH}}^{\text{теор}} \cdot \frac{k_{\text{изб}}}{\eta_{\text{скр}}}\quad (\text{кг/ч})$$
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

