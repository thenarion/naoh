"""
Модуль инженерных и химических расчетов материального баланса и расхода чистого 100% NaOH
для системы газоочистки (мокрого скруббера) термических установок.

Расчеты разделены для двух автономных установок:
1) Установка утилизации жидких отходов (КТОЖС, расход Q = 1.5 м³/ч)
2) Установка утилизации ТБО (расход M = 170 кг/ч)
А также поддерживается сводный (комбинированный) расчет.

В расчет включены ТОЛЬКО активные кислотообразующие компоненты:
- Хлор (Cl⁻ в жидких стоках, Cl в ТБО) -> образует HCl
- Сера (SO₄²⁻ в жидких стоках, S в ТБО) -> образует SO₂

Расчет ведется строго на 100% чистый реагент (NaOH).
"""

from typing import Dict, Any, List, Tuple, Optional
import numpy as np

# ==========================================
# 1. МОЛЯРНЫЕ МАССЫ И ХИМИЧЕСКИЕ КОНСТАНТЫ
# ==========================================
M_H = 1.008       # г/моль
M_O = 15.999      # г/моль
M_Na = 22.990     # г/моль
M_S = 32.065      # г/моль
M_Cl = 35.453     # г/моль

# Молярные массы соединений
M_HCl = M_H + M_Cl                  # 36.461 г/моль
M_NaOH = M_Na + M_O + M_H           # 39.997 г/моль (~40.0)
M_SO2 = M_S + 2 * M_O               # 64.063 г/моль (~64.0)
M_SO4 = M_S + 4 * M_O               # 96.061 г/моль (~96.0)
M_NaCl = M_Na + M_Cl                # 58.443 г/моль
M_Na2SO3 = 2 * M_Na + M_S + 3 * M_O # 126.04 г/моль

# ==========================================
# 2. ИСХОДНЫЕ НАБОРЫ ДАННЫХ ПО ЖИДКИМ ОТХОДАМ
# ==========================================
# Согласно Таблице Г.1 СП 320.1325800.2017
LIQUID_WASTE_DATASETS = {
    "young": {
        "id": "young",
        "name": "Набор 1: Молодой полигон (кислая фаза)",
        "phase": "Кислая фаза",
        "norm_ref": "СП 320.1325800.2017, Таблица Г.1",
        "description": "Период активного брожения и выщелачивания. Пиковые проектные концентрации.",
        "cl_range": (300.0, 2650.0, 5000.0),    # (min, avg, max), мг/дм³
        "so4_range": (40.0, 770.0, 1500.0),     # (min, avg, max), мг/дм³
        "default_cl": 5000.0,
        "default_so4": 1500.0
    },
    "old": {
        "id": "old",
        "name": "Набор 2: Старый полигон (метаногенная фаза)",
        "phase": "Метаногенная фаза",
        "norm_ref": "СП 320.1325800.2017, Таблица Г.1",
        "description": "Стабилизированная метаногенная фаза. Умеренные концентрации.",
        "cl_range": (300.0, 1400.0, 2500.0),    # (min, avg, max), мг/дм³
        "so4_range": (25.0, 212.5, 400.0),      # (min, avg, max), мг/дм³
        "default_cl": 2500.0,
        "default_so4": 400.0
    }
}

# ==========================================
# 3. ИСХОДНЫЕ ДАННЫЕ И СОСТАВ ТБО (Cl и S)
# ==========================================
TBO_WASTE_GROUPS: Dict[str, Dict[str, Any]] = {
    "Нефтепродукты, масла, шламы": {
        "cl": 0.02, "s": 0.50,
        "description": "Всплывшие нефтепродукты, отработанные масла, шлам очистки емкостей, остатки дизтоплива"
    },
    "Резина, РТИ, обувь": {
        "cl": 0.50, "s": 1.00,
        "description": "Обрезки вулканизованной резины, резиновая обувь, сальниковая набивка, РТИ"
    },
    "Хлорсодержащие полимеры (ПВХ, хлоралканы)": {
        "cl": 3.50, "s": 0.10,
        "description": "Обводненные хлоралканы, полимерная тара спецназначения, ПВХ-материалы"
    },
    "Общие полимеры (ПЭ, ПП, ПС, полиуретан)": {
        "cl": 0.10, "s": 0.05,
        "description": "Лом полистирола, полиуретан, тара полиэтиленовая, синтетические волокна"
    },
    "Древесина, бумага, картон": {
        "cl": 0.05, "s": 0.05,
        "description": "Тара деревянная, опилки, фильтры бумажные, отходы упаковки из картона"
    },
    "Пищевые отходы кухонь": {
        "cl": 0.30, "s": 0.15,
        "description": "Пищевые отходы кухонь и столовых несортированные"
    },
    "Текстиль, спецодежда, ветошь": {
        "cl": 0.25, "s": 0.20,
        "description": "Спецодежда х/б и шерсть, обтирочный материал загрязненный, нетканые материалы"
    },
    "Антифризы и теплоносители": {
        "cl": 0.10, "s": 0.10,
        "description": "Отходы антифризов на основе этиленгликоля и пропиленгликоля"
    },
    "Отработанные фильтры и сорбенты": {
        "cl": 0.15, "s": 0.30,
        "description": "Фильтры масляные/топливные/воздушные, сорбенты на основе торфа, активированный уголь"
    },
    "Прочие смешанные отходы и мусор": {
        "cl": 0.20, "s": 0.15,
        "description": "Мусор от офисных и бытовых помещений, несортированные отходы"
    }
}

TBO_PRESETS: Dict[str, Dict[str, float]] = {
    "Усредненный смешанный состав (рекомендуемый)": {
        "Нефтепродукты, масла, шламы": 12.0,
        "Резина, РТИ, обувь": 6.0,
        "Хлорсодержащие полимеры (ПВХ, хлоралканы)": 3.0,
        "Общие полимеры (ПЭ, ПП, ПС, полиуретан)": 15.0,
        "Древесина, бумага, картон": 20.0,
        "Пищевые отходы кухонь": 10.0,
        "Текстиль, спецодежда, ветошь": 10.0,
        "Антифризы и теплоносители": 5.0,
        "Отработанные фильтры и сорбенты": 9.0,
        "Прочие смешанные отходы и мусор": 10.0,
    },
    "Период строительства и обустройства": {
        "Нефтепродукты, масла, шламы": 8.0,
        "Резина, РТИ, обувь": 8.0,
        "Хлорсодержащие полимеры (ПВХ, хлоралканы)": 2.0,
        "Общие полимеры (ПЭ, ПП, ПС, полиуретан)": 12.0,
        "Древесина, бумага, картон": 28.0,
        "Пищевые отходы кухонь": 8.0,
        "Текстиль, спецодежда, ветошь": 12.0,
        "Антифризы и теплоносители": 5.0,
        "Отработанные фильтры и сорбенты": 10.0,
        "Прочие смешанные отходы и мусор": 7.0,
    },
    "Период промышленной эксплуатации скважин и ДКС": {
        "Нефтепродукты, масла, шламы": 22.0,
        "Резина, РТИ, обувь": 6.0,
        "Хлорсодержащие полимеры (ПВХ, хлоралканы)": 5.0,
        "Общие полимеры (ПЭ, ПП, ПС, полиуретан)": 12.0,
        "Древесина, бумага, картон": 10.0,
        "Пищевые отходы кухонь": 5.0,
        "Текстиль, спецодежда, ветошь": 10.0,
        "Антифризы и теплоносители": 10.0,
        "Отработанные фильтры и сорбенты": 14.0,
        "Прочие смешанные отходы и мусор": 6.0,
    }
}


# ==========================================
# 4. РАСЧЕТ ПОСТУПЛЕНИЯ ГЕТЕРОАТОМОВ
# ==========================================

def calculate_liquid_waste_pollutants(
    q_liq: float,
    c_cl: float,
    c_so4: float,
    k_conv_cl: float = 0.95,
    k_conv_s: float = 0.85,
    dataset_name: str = "Пользовательский"
) -> Dict[str, Any]:
    """
    Расчет поступления Cl, S и выхода газов HCl, SO2 из жидких отходов.

    Научное обоснование (при Т дожигателя = 1100 °C):
    - Для жидких отходов (распыление в факел): конверсия Cl→HCl максимальна (0.95),
      но часть S (до 15%) связывается в золе в виде сульфатов (k_conv_s = 0.85).
    """
    mass_cl = (q_liq * c_cl) / 1000.0  # кг/ч
    sulfate_to_s_ratio = M_S / M_SO4     # 32.065 / 96.061 ~ 0.3338
    mass_s = (q_liq * c_so4 * sulfate_to_s_ratio) / 1000.0 # кг/ч
    
    mass_hcl = mass_cl * k_conv_cl * (M_HCl / M_Cl)
    mass_so2 = mass_s * k_conv_s * (M_SO2 / M_S)
    
    feed_mass_kg_h = q_liq * 1000.0  # кг/ч при плотности ~1000 кг/м³
    
    return {
        "stream_type": "liquid",
        "dataset_name": dataset_name,
        "q_liq": q_liq,
        "feed_mass_kg_h": feed_mass_kg_h,
        "c_cl": c_cl,
        "c_so4": c_so4,
        "mass_cl": mass_cl,
        "mass_s": mass_s,
        "mass_hcl": mass_hcl,
        "mass_so2": mass_so2,
        "k_conv_cl": k_conv_cl,
        "k_conv_s": k_conv_s
    }


def calculate_tbo_pollutants(
    m_tbo: float,
    morphology_dict: Dict[str, float] = None,
    custom_elements: Dict[str, float] = None,
    k_conv_cl: float = 0.85,
    k_conv_s: float = 0.80
) -> Dict[str, Any]:
    """
    Расчет поступления Cl, S и выхода газов HCl, SO2 из ТБО.

    Научное обоснование (при Т дожигателя = 1100 °C):
    - Для ТБО (слоевое сжигание): значительная часть Cl (до 15%) связывается в золе 
      в виде NaCl/KCl (k_conv_cl = 0.85), а до 20% S — в виде CaSO₄ (k_conv_s = 0.80).
    """
    breakdown_table = []
    
    if custom_elements is not None:
        pct_cl = custom_elements.get("cl", 0.35)
        pct_s = custom_elements.get("s", 0.30)
        
        total_cl = m_tbo * (pct_cl / 100.0)
        total_s = m_tbo * (pct_s / 100.0)
        
        breakdown_table.append({
            "Группа отходов": "Смешанные ТБО (прямой элементный состав)",
            "Доля, %": 100.0,
            "Масса отхода, кг/ч": m_tbo,
            "Cl в отходе, кг/ч": round(total_cl, 4),
            "S в отходе, кг/ч": round(total_s, 4),
            "Cl, %": pct_cl,
            "S, %": pct_s
        })
    else:
        if morphology_dict is None:
            morphology_dict = TBO_PRESETS["Усредненный смешанный состав (рекомендуемый)"]
            
        total_cl = 0.0
        total_s = 0.0
        
        for group_name, share_pct in morphology_dict.items():
            if share_pct <= 0:
                continue
            props = TBO_WASTE_GROUPS.get(group_name, {"cl": 0.20, "s": 0.20})
            m_group = m_tbo * (share_pct / 100.0)
            
            m_cl_g = m_group * (props["cl"] / 100.0)
            m_s_g = m_group * (props["s"] / 100.0)
            
            total_cl += m_cl_g
            total_s += m_s_g
            
            breakdown_table.append({
                "Группа отходов": group_name,
                "Доля, %": share_pct,
                "Масса отхода, кг/ч": round(m_group, 2),
                "Cl в отходе, кг/ч": round(m_cl_g, 4),
                "S в отходе, кг/ч": round(m_s_g, 4),
                "Cl, %": props["cl"],
                "S, %": props["s"]
            })
    
    mass_hcl = total_cl * k_conv_cl * (M_HCl / M_Cl)
    mass_so2 = total_s * k_conv_s * (M_SO2 / M_S)
    
    avg_pct_cl = (total_cl / m_tbo * 100.0) if m_tbo > 0 else 0.0
    avg_pct_s = (total_s / m_tbo * 100.0) if m_tbo > 0 else 0.0
    
    return {
        "stream_type": "tbo",
        "m_tbo": m_tbo,
        "feed_mass_kg_h": m_tbo,
        "mass_cl": total_cl,
        "mass_s": total_s,
        "avg_pct_cl": avg_pct_cl,
        "avg_pct_s": avg_pct_s,
        "mass_hcl": mass_hcl,
        "mass_so2": mass_so2,
        "breakdown_table": breakdown_table,
        "k_conv_cl": k_conv_cl,
        "k_conv_s": k_conv_s
    }




# ==========================================
# 5. РАСЧЕТ РАСХОДА 100% NaOH ДЛЯ ДОСТИЖЕНИЯ ПДК ИТС 9-2020
# ==========================================

def calculate_single_stream_consumption(
    mass_hcl_gas: float,
    mass_so2_gas: float,
    feed_cl: float,
    feed_s: float,
    feed_mass_kg_h: float,
    k_conv_cl: float = 1.0,
    k_conv_s: float = 1.0,
    flue_gas_flow: float = 3000.0,
    k_excess: float = 1.15,
    limit_hcl: float = 10.0,
    limit_so2: float = 50.0,
    eta_scrubber: float = 0.95,
    hours_per_day: float = 24.0,
    operating_days_year: float = 365.0,
    title: str = "Установка"
) -> Dict[str, Any]:
    """
    Расчет расхода 100% чистого NaOH для достижения нормативов ИТС 9-2020
    (HCl ≤ 10 мг/нм³, SO₂ ≤ 50 мг/нм³) в зависимости от расхода дымовых газов flue_gas_flow (нм³/ч).
    """
    # 1. Входные концентрации в дымовых газах (мг/нм³)
    conc_hcl_in = (mass_hcl_gas * 1e6) / flue_gas_flow if flue_gas_flow > 0 else 0.0
    conc_so2_in = (mass_so2_gas * 1e6) / flue_gas_flow if flue_gas_flow > 0 else 0.0
    
    # 2. Требуемая степень очистки для выхода на предельные значения ИТС 9-2020
    eta_req_hcl = max(0.0, 1.0 - (limit_hcl / conc_hcl_in)) if conc_hcl_in > limit_hcl else 0.0
    eta_req_so2 = max(0.0, 1.0 - (limit_so2 / conc_so2_in)) if conc_so2_in > limit_so2 else 0.0
    
    # 3. Масса кислых газов, улавливаемая в скруббере при паспортной эффективности eta_scrubber (кг/ч)
    mass_hcl_neut = mass_hcl_gas * eta_scrubber
    mass_so2_neut = mass_so2_gas * eta_scrubber
    
    # Фактические выходные концентрации в очищенных газах после скруббера (мг/нм³)
    conc_hcl_out = conc_hcl_in * (1.0 - eta_scrubber)
    conc_so2_out = conc_so2_in * (1.0 - eta_scrubber)

    
    # 4. Стехиометрические коэффициенты
    stoich_hcl = M_NaOH / M_HCl         # ~1.09697 кг NaOH / кг HCl
    stoich_so2 = (2 * M_NaOH) / M_SO2   # ~1.24867 кг NaOH / кг SO2
    
    # 5. Теоретический расход 100% NaOH на нейтрализацию уловленных кислых газов (кг/ч)
    naoh_hcl_theor = mass_hcl_neut * stoich_hcl
    naoh_so2_theor = mass_so2_neut * stoich_so2
    naoh_total_theor = naoh_hcl_theor + naoh_so2_theor
    
    # 6. Фактический часовой расход 100% NaOH с учетом технологического избытка k_изб (кг/ч)
    naoh_hcl_fact = naoh_hcl_theor * k_excess
    naoh_so2_fact = naoh_so2_theor * k_excess
    naoh_total_fact = naoh_hcl_fact + naoh_so2_fact
    
    # 7. Режим работы: часы в сутки и дни в году
    operating_hours_year = hours_per_day * operating_days_year
    
    # 8. Суточный и годовой расход 100% чистого NaOH
    naoh_pure_day_kg = naoh_total_fact * hours_per_day
    naoh_pure_year_t = (naoh_total_fact * operating_hours_year) / 1000.0
    
    # 9. Удельный расход на 1 кг отхода
    if feed_mass_kg_h > 0:
        spec_naoh_pure_kg_per_kg = naoh_total_fact / feed_mass_kg_h       # кг 100% NaOH / кг отхода
        spec_naoh_pure_g_per_kg = spec_naoh_pure_kg_per_kg * 1000.0       # г 100% NaOH / кг отхода
    else:
        spec_naoh_pure_kg_per_kg = 0.0
        spec_naoh_pure_g_per_kg = 0.0
        
    return {
        "title": title,
        "flue_gas_flow": flue_gas_flow,
        "feed_mass_kg_h": feed_mass_kg_h,
        "feed_cl": feed_cl,
        "feed_s": feed_s,
        "k_conv_cl": k_conv_cl,
        "k_conv_s": k_conv_s,
        
        "mass_hcl_gas": mass_hcl_gas,
        "mass_so2_gas": mass_so2_gas,
        "mass_hcl_neut": mass_hcl_neut,
        "mass_so2_neut": mass_so2_neut,
        
        "conc_hcl_in": conc_hcl_in,
        "conc_so2_in": conc_so2_in,
        "conc_hcl_out": conc_hcl_out,
        "conc_so2_out": conc_so2_out,
        "limit_hcl": limit_hcl,
        "limit_so2": limit_so2,
        "eta_req_hcl": eta_req_hcl,
        "eta_req_so2": eta_req_so2,
        
        "stoich_hcl": stoich_hcl,
        "stoich_so2": stoich_so2,
        
        "naoh_hcl_theor": naoh_hcl_theor,
        "naoh_so2_theor": naoh_so2_theor,
        "naoh_total_theor": naoh_total_theor,
        
        "naoh_hcl_fact": naoh_hcl_fact,
        "naoh_so2_fact": naoh_so2_fact,
        "naoh_total_fact": naoh_total_fact,
        
        # Режим работы
        "hours_per_day": hours_per_day,
        "operating_days_year": operating_days_year,
        "operating_hours_year": operating_hours_year,
        
        # 100% чистый реагент
        "naoh_pure_hour_kg": naoh_total_fact,
        "naoh_pure_day_kg": naoh_pure_day_kg,
        "naoh_pure_year_t": naoh_pure_year_t,
        
        # УДЕЛЬНЫЙ РАСХОД НА 1 КГ ОТХОДА
        "spec_naoh_pure_kg_per_kg": spec_naoh_pure_kg_per_kg,
        "spec_naoh_pure_g_per_kg": spec_naoh_pure_g_per_kg,
        
        "eta_scrubber": eta_scrubber,
        "k_excess": k_excess
    }


def calculate_liquid_installation_naoh(
    liquid_results: Dict[str, Any],
    flue_gas_flow: float = 2500.0,
    k_excess: float = 1.15,
    eta_scrubber: float = 0.95,
    hours_per_day: float = 24.0,
    operating_days_year: float = 365.0
) -> Dict[str, Any]:
    """
    Отдельный расчет расхода NaOH для Установки утилизации жидких отходов (КТОЖС, 1,5 м³/ч)
    на соблюдение нормативов ИТС 9-2020 при собственном расходе дымовых газов flue_gas_flow (нм³/ч).
    """
    feed_mass = liquid_results.get("feed_mass_kg_h", liquid_results.get("q_liq", 1.5) * 1000.0)
    res = calculate_single_stream_consumption(
        mass_hcl_gas=liquid_results["mass_hcl"],
        mass_so2_gas=liquid_results["mass_so2"],
        feed_cl=liquid_results["mass_cl"],
        feed_s=liquid_results["mass_s"],
        feed_mass_kg_h=feed_mass,
        k_conv_cl=liquid_results.get("k_conv_cl", 0.95),
        k_conv_s=liquid_results.get("k_conv_s", 0.85),
        flue_gas_flow=flue_gas_flow,
        k_excess=k_excess,
        eta_scrubber=eta_scrubber,
        hours_per_day=hours_per_day,
        operating_days_year=operating_days_year,
        title="Установка утилизации жидких отходов"
    )
    res["q_liq"] = liquid_results["q_liq"]
    res["spec_naoh_pure_per_m3_kg"] = res["naoh_pure_hour_kg"] / liquid_results["q_liq"] if liquid_results["q_liq"] > 0 else 0.0
    return res


def calculate_tbo_installation_naoh(
    tbo_results: Dict[str, Any],
    flue_gas_flow: float = 800.0,
    k_excess: float = 1.15,
    eta_scrubber: float = 0.95,
    hours_per_day: float = 24.0,
    operating_days_year: float = 365.0
) -> Dict[str, Any]:
    """
    Отдельный расчет расхода NaOH для Установки утилизации ТБО (170 кг/ч)
    на соблюдение нормативов ИТС 9-2020 при собственном расходе дымовых газов flue_gas_flow (нм³/ч).
    """
    feed_mass = tbo_results.get("feed_mass_kg_h", tbo_results.get("m_tbo", 170.0))
    res = calculate_single_stream_consumption(
        mass_hcl_gas=tbo_results["mass_hcl"],
        mass_so2_gas=tbo_results["mass_so2"],
        feed_cl=tbo_results["mass_cl"],
        feed_s=tbo_results["mass_s"],
        feed_mass_kg_h=feed_mass,
        k_conv_cl=tbo_results.get("k_conv_cl", 0.85),
        k_conv_s=tbo_results.get("k_conv_s", 0.80),
        flue_gas_flow=flue_gas_flow,
        k_excess=k_excess,
        eta_scrubber=eta_scrubber,
        hours_per_day=hours_per_day,
        operating_days_year=operating_days_year,
        title="Установка утилизации ТБО"
    )
    res["m_tbo"] = tbo_results["m_tbo"]
    return res



def calculate_combined_installations_naoh(
    liquid_res_calc: Dict[str, Any],
    tbo_res_calc: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Сводный расчет при совместной работе обеих автономных установок с раздельными потоками газов.
    """
    flue_gas_flow_liq = liquid_res_calc.get("flue_gas_flow", 2500.0)
    flue_gas_flow_tbo = tbo_res_calc.get("flue_gas_flow", 800.0)
    flue_gas_flow_total = flue_gas_flow_liq + flue_gas_flow_tbo
    
    pure_hour = liquid_res_calc["naoh_pure_hour_kg"] + tbo_res_calc["naoh_pure_hour_kg"]
    pure_day = liquid_res_calc["naoh_pure_day_kg"] + tbo_res_calc["naoh_pure_day_kg"]
    pure_year = liquid_res_calc["naoh_pure_year_t"] + tbo_res_calc["naoh_pure_year_t"]
    
    total_feed_mass_kg_h = liquid_res_calc.get("feed_mass_kg_h", 0.0) + tbo_res_calc.get("feed_mass_kg_h", 0.0)
    
    if total_feed_mass_kg_h > 0:
        spec_naoh_pure_kg_per_kg = pure_hour / total_feed_mass_kg_h
        spec_naoh_pure_g_per_kg = spec_naoh_pure_kg_per_kg * 1000.0
    else:
        spec_naoh_pure_kg_per_kg = 0.0
        spec_naoh_pure_g_per_kg = 0.0
        
    return {
        "title": "Суммарно по двум установкам",
        "flue_gas_flow_liq": flue_gas_flow_liq,
        "flue_gas_flow_tbo": flue_gas_flow_tbo,
        "flue_gas_flow": flue_gas_flow_total,
        "total_feed_mass_kg_h": total_feed_mass_kg_h,
        "naoh_pure_hour_kg": pure_hour,
        "naoh_pure_day_kg": pure_day,
        "naoh_pure_year_t": pure_year,
        "naoh_total_theor": liquid_res_calc.get("naoh_total_theor", 0.0) + tbo_res_calc.get("naoh_total_theor", 0.0),
        "naoh_total_fact": pure_hour,
        "spec_naoh_pure_kg_per_kg": spec_naoh_pure_kg_per_kg,
        "spec_naoh_pure_g_per_kg": spec_naoh_pure_g_per_kg
    }


def calculate_naoh_and_compliance(
    liquid_results: Dict[str, Any],
    tbo_results: Dict[str, Any],
    flue_gas_flow_liq: float = 2500.0,
    flue_gas_flow_tbo: float = 800.0,
    flue_gas_flow: Optional[float] = None,
    eta_scrubber: float = 0.95,
    k_excess: float = 1.15,
    eta_co2_abs: float = 0.0,
    c_naoh_sol: float = 100.0,
    hours_per_day: float = 24.0,
    operating_days_year: float = 365.0
) -> Dict[str, Any]:
    """
    Расчет расхода NaOH на достижение нормативов выбросов ИТС 9-2020 (Приложение В)
    для двух автономных установок с раздельными расходами дымовых газов.
    """
    if flue_gas_flow is not None and flue_gas_flow_liq == 2500.0 and flue_gas_flow_tbo == 800.0:
        # Для обратной совместимости, если передан один суммарный поток
        feed_l = liquid_results.get('feed_mass_kg_h', 1500.0)
        feed_t = tbo_results.get('feed_mass_kg_h', 170.0)
        tot_f = feed_l + feed_t
        flue_gas_flow_liq = flue_gas_flow * (feed_l / tot_f) if tot_f > 0 else flue_gas_flow * 0.9
        flue_gas_flow_tbo = flue_gas_flow - flue_gas_flow_liq
        
    flue_gas_flow_total = flue_gas_flow_liq + flue_gas_flow_tbo

    # 1. Расчет для жидких отходов
    mass_hcl_liq = liquid_results.get('mass_hcl', 0.0)
    mass_so2_liq = liquid_results.get('mass_so2', 0.0)
    conc_hcl_in_liq = (mass_hcl_liq * 1e6) / flue_gas_flow_liq if flue_gas_flow_liq > 0 else 0.0
    conc_so2_in_liq = (mass_so2_liq * 1e6) / flue_gas_flow_liq if flue_gas_flow_liq > 0 else 0.0
    eta_req_hcl_liq = max(0.0, 1.0 - (10.0 / conc_hcl_in_liq)) if conc_hcl_in_liq > 10.0 else 0.0
    eta_req_so2_liq = max(0.0, 1.0 - (50.0 / conc_so2_in_liq)) if conc_so2_in_liq > 50.0 else 0.0

    mass_hcl_neut_liq = mass_hcl_liq * eta_scrubber
    mass_so2_neut_liq = mass_so2_liq * eta_scrubber


    # 2. Расчет для ТБО
    mass_hcl_tbo = tbo_results.get('mass_hcl', 0.0)
    mass_so2_tbo = tbo_results.get('mass_so2', 0.0)
    conc_hcl_in_tbo = (mass_hcl_tbo * 1e6) / flue_gas_flow_tbo if flue_gas_flow_tbo > 0 else 0.0
    conc_so2_in_tbo = (mass_so2_tbo * 1e6) / flue_gas_flow_tbo if flue_gas_flow_tbo > 0 else 0.0
    eta_req_hcl_tbo = max(0.0, 1.0 - (10.0 / conc_hcl_in_tbo)) if conc_hcl_in_tbo > 10.0 else 0.0
    eta_req_so2_tbo = max(0.0, 1.0 - (50.0 / conc_so2_in_tbo)) if conc_so2_in_tbo > 50.0 else 0.0
    mass_hcl_neut_tbo = mass_hcl_tbo * eta_scrubber
    mass_so2_neut_tbo = mass_so2_tbo * eta_scrubber

    # 3. Суммарные потоки
    mass_hcl_in_total = mass_hcl_liq + mass_hcl_tbo
    mass_so2_in_total = mass_so2_liq + mass_so2_tbo
    mass_hcl_neut_total = mass_hcl_neut_liq + mass_hcl_neut_tbo
    mass_so2_neut_total = mass_so2_neut_liq + mass_so2_neut_tbo
    mass_co2_abs = tbo_results.get('mass_co2', 0.0) * eta_co2_abs

    # Взвешенные концентрации в смеси дымовых газов
    conc_hcl_in_total = (mass_hcl_in_total * 1e6) / flue_gas_flow_total if flue_gas_flow_total > 0 else 0.0
    conc_so2_in_total = (mass_so2_in_total * 1e6) / flue_gas_flow_total if flue_gas_flow_total > 0 else 0.0
    eta_req_hcl_total = (mass_hcl_neut_total / mass_hcl_in_total) if mass_hcl_in_total > 0 else 0.0
    eta_req_so2_total = (mass_so2_neut_total / mass_so2_in_total) if mass_so2_in_total > 0 else 0.0

    conc_hcl_out_total = conc_hcl_in_total * (1.0 - eta_scrubber)
    conc_so2_out_total = conc_so2_in_total * (1.0 - eta_scrubber)


    # 4. Стехиометрический и фактический расход 100% NaOH
    stoich_hcl = M_NaOH / M_HCl         # ~1.09697 кг NaOH / кг HCl
    stoich_so2 = (2 * M_NaOH) / M_SO2   # ~1.24867 кг NaOH / кг SO2
    
    naoh_hcl_theor = mass_hcl_neut_total * stoich_hcl
    naoh_so2_theor = mass_so2_neut_total * stoich_so2
    naoh_total_theor = naoh_hcl_theor + naoh_so2_theor
    
    naoh_hcl_fact = naoh_hcl_theor * k_excess
    naoh_so2_fact = naoh_so2_theor * k_excess
    naoh_total_fact = naoh_hcl_fact + naoh_so2_fact
    
    operating_hours_year = hours_per_day * operating_days_year
    naoh_pure_day_kg = naoh_total_fact * hours_per_day
    naoh_pure_year_t = (naoh_total_fact * operating_hours_year) / 1000.0
    
    total_feed_mass_kg_h = liquid_results.get('feed_mass_kg_h', 0.0) + tbo_results.get('feed_mass_kg_h', 0.0)
    spec_naoh_pure_kg_per_kg = (naoh_total_fact / total_feed_mass_kg_h) if total_feed_mass_kg_h > 0 else 0.0
    spec_naoh_pure_g_per_kg = spec_naoh_pure_kg_per_kg * 1000.0
    
    # 5. Проверка достаточности скрубберов
    compliance_status = "✅ НОРМА"
    compliance_msg = ""
    max_req_eta = max(eta_req_hcl_liq, eta_req_so2_liq, eta_req_hcl_tbo, eta_req_so2_tbo)
    if eta_scrubber < max_req_eta:
        compliance_status = "⚠️ ТРЕБУЕТСЯ ВНИМАНИЕ"
        problems = []
        if eta_scrubber < eta_req_hcl_liq:
            problems.append(f"Уст.1 (Жидкие) по HCl: η_req = {eta_req_hcl_liq:.1%} (при V_г = {flue_gas_flow_liq:.0f} нм³/ч)")
        if eta_scrubber < eta_req_so2_liq:
            problems.append(f"Уст.1 (Жидкие) по SO₂: η_req = {eta_req_so2_liq:.1%} (при V_г = {flue_gas_flow_liq:.0f} нм³/ч)")
        if eta_scrubber < eta_req_hcl_tbo:
            problems.append(f"Уст.2 (ТБО) по HCl: η_req = {eta_req_hcl_tbo:.1%} (при V_г = {flue_gas_flow_tbo:.0f} нм³/ч)")
        if eta_scrubber < eta_req_so2_tbo:
            problems.append(f"Уст.2 (ТБО) по SO₂: η_req = {eta_req_so2_tbo:.1%} (при V_г = {flue_gas_flow_tbo:.0f} нм³/ч)")
        compliance_msg = f"Для выхода на нормативы ИТС 9-2020 при паспортном КПД скруббера {eta_scrubber:.1%}: {'; '.join(problems)}."
    
    return {
        'title': "Суммарно по комплексу с расчетом под ИТС 9-2020",
        'flue_gas_flow_liq': flue_gas_flow_liq,
        'flue_gas_flow_tbo': flue_gas_flow_tbo,
        'flue_gas_flow': flue_gas_flow_total,
        
        'mass_hcl_liq': mass_hcl_liq,
        'mass_so2_liq': mass_so2_liq,
        'conc_hcl_in_liq': conc_hcl_in_liq,
        'conc_so2_in_liq': conc_so2_in_liq,
        'eta_req_hcl_liq': eta_req_hcl_liq,
        'eta_req_so2_liq': eta_req_so2_liq,
        
        'mass_hcl_tbo': mass_hcl_tbo,
        'mass_so2_tbo': mass_so2_tbo,
        'conc_hcl_in_tbo': conc_hcl_in_tbo,
        'conc_so2_in_tbo': conc_so2_in_tbo,
        'eta_req_hcl_tbo': eta_req_hcl_tbo,
        'eta_req_so2_tbo': eta_req_so2_tbo,
        
        'mass_hcl_in': mass_hcl_in_total,
        'mass_so2_in': mass_so2_in_total,
        'mass_hcl_neut': mass_hcl_neut_total,
        'mass_so2_neut': mass_so2_neut_total,
        'mass_co2_abs': mass_co2_abs,
        'conc_hcl_in': conc_hcl_in_total,
        'conc_so2_in': conc_so2_in_total,
        'conc_hcl_out': conc_hcl_out_total,
        'conc_so2_out': conc_so2_out_total,
        'limit_hcl': 10.0,
        'limit_so2': 50.0,
        'eta_req_hcl': eta_req_hcl_total,
        'eta_req_so2': eta_req_so2_total,
        'compliance_status': compliance_status,
        'compliance_msg': compliance_msg,
        
        'total_feed_mass_kg_h': total_feed_mass_kg_h,
        'stoich_hcl': stoich_hcl,
        'stoich_so2': stoich_so2,
        'naoh_hcl_theor': naoh_hcl_theor,
        'naoh_so2_theor': naoh_so2_theor,
        'naoh_total_theor': naoh_total_theor,
        'naoh_hcl_fact': naoh_hcl_fact,
        'naoh_so2_fact': naoh_so2_fact,
        'naoh_total_fact': naoh_total_fact,
        
        'hours_per_day': hours_per_day,
        'operating_days_year': operating_days_year,
        'operating_hours_year': operating_hours_year,
        
        'naoh_pure_hour_kg': naoh_total_fact,
        'naoh_pure_day_kg': naoh_pure_day_kg,
        'naoh_pure_year_t': naoh_pure_year_t,
        'spec_naoh_pure_kg_per_kg': spec_naoh_pure_kg_per_kg,
        'spec_naoh_pure_g_per_kg': spec_naoh_pure_g_per_kg,
        'eta_scrubber': eta_scrubber,
        'k_excess': k_excess
    }
    
    total_feed_mass_kg_h = liquid_results.get('feed_mass_kg_h', 0.0) + tbo_results.get('feed_mass_kg_h', 0.0)
    spec_naoh_pure_kg_per_kg = (naoh_total_fact / total_feed_mass_kg_h) if total_feed_mass_kg_h > 0 else 0.0
    spec_naoh_pure_g_per_kg = spec_naoh_pure_kg_per_kg * 1000.0
    
    # 6. Проверка технической достаточности скруббера с КПД eta_scrubber
    compliance_status = "✅ НОРМА"
    compliance_msg = ""
    if eta_scrubber < eta_req_hcl or eta_scrubber < eta_req_so2:
        compliance_status = "⚠️ ТРЕБУЕТСЯ ВНИМАНИЕ"
        problems = []
        if eta_scrubber < eta_req_hcl:
            problems.append(f"по HCl требуется η ≥ {eta_req_hcl:.1%} (паспортная эффективность: {eta_scrubber:.1%})")
        if eta_scrubber < eta_req_so2:
            problems.append(f"по SO₂ требуется η ≥ {eta_req_so2:.1%} (паспортная эффективность: {eta_scrubber:.1%})")
        compliance_msg = f"Для выхода на нормативы ИТС 9-2020 при расходе газов {flue_gas_flow:.0f} нм³/ч: {'; '.join(problems)}."
    
    return {
        'title': "Суммарно по комплексу с расчетом под ИТС 9-2020",
        'flue_gas_flow': flue_gas_flow,
        'mass_hcl_in': mass_hcl_in,
        'mass_so2_in': mass_so2_in,
        'mass_hcl_neut': mass_hcl_neut,
        'mass_so2_neut': mass_so2_neut,
        'mass_co2_abs': mass_co2_abs,
        'conc_hcl_in': conc_hcl_in,
        'conc_so2_in': conc_so2_in,
        'conc_hcl_out': conc_hcl_out,
        'conc_so2_out': conc_so2_out,
        'limit_hcl': 10.0,
        'limit_so2': 50.0,
        'eta_req_hcl': eta_req_hcl,
        'eta_req_so2': eta_req_so2,
        'compliance_status': compliance_status,
        'compliance_msg': compliance_msg,
        
        'total_feed_mass_kg_h': total_feed_mass_kg_h,
        'stoich_hcl': stoich_hcl,
        'stoich_so2': stoich_so2,
        'naoh_hcl_theor': naoh_hcl_theor,
        'naoh_so2_theor': naoh_so2_theor,
        'naoh_total_theor': naoh_total_theor,
        'naoh_hcl_fact': naoh_hcl_fact,
        'naoh_so2_fact': naoh_so2_fact,
        'naoh_total_fact': naoh_total_fact,
        
        'hours_per_day': hours_per_day,
        'operating_days_year': operating_days_year,
        'operating_hours_year': operating_hours_year,
        
        'naoh_pure_hour_kg': naoh_total_fact,
        'naoh_pure_day_kg': naoh_pure_day_kg,
        'naoh_pure_year_t': naoh_pure_year_t,
        'spec_naoh_pure_kg_per_kg': spec_naoh_pure_kg_per_kg,
        'spec_naoh_pure_g_per_kg': spec_naoh_pure_g_per_kg,
        'eta_scrubber': eta_scrubber,
        'k_excess': k_excess
    }

