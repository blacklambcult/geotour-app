import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
import requests
import io
import math
import pandas as pd
from fpdf import FPDF

st.set_page_config(page_title="GeoTour: 2GIS Геоаналитика", layout="wide")

st.title("🏨 GeoTour: Геомаркетинговый аудит и Конструктор туров")
st.caption("Система поддержки маркетинговых решений туристского предприятия на базе 2GIS Places API (Республика Татарстан)")

# API Ключ 2ГИС
TWOGIS_API_KEY = "a7f48cf3-379c-43cb-8f31-99a8688f1319"

# 1. Боковая панель
st.sidebar.header("Параметры анализа")
radius_km = st.sidebar.slider("Радиус сканирования (км)", 1, 20, 5, step=1)
radius_m = radius_km * 1000

st.sidebar.subheader("Слои карты аудита")
show_poi = st.sidebar.checkbox("Туристские магниты (POI)", value=True)
show_comp = st.sidebar.checkbox("Конкуренты (Отели/Базы)", value=True)
show_infra = st.sidebar.checkbox("Инфраструктура (Кафе/Рестораны)", value=True)

# 2. Математические расчеты
def haversine_distance(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def solve_tsp_nearest_neighbor(start_coords, points):
    if not points:
        return []
    unvisited = points.copy()
    route = []
    cur_lat, cur_lon = start_coords
    while unvisited:
        nearest_idx = 0
        min_dist = float('inf')
        for i, pt in enumerate(unvisited):
            d = haversine_distance(cur_lat, cur_lon, pt['lat'], pt['lon'])
            if d < min_dist:
                min_dist = d
                nearest_idx = i
        next_pt = unvisited.pop(nearest_idx)
        route.append(next_pt)
        cur_lat, cur_lon = next_pt['lat'], next_pt['lon']
    return route

# 3. Запрос данных к 2GIS Places API
@st.cache_data(show_spinner=False, ttl=1800)
def fetch_2gis_places(lat, lon, radius_meters, query_text, cat_name):
    url = "https://catalog.api.2gis.com/3.0/items"
    
    # 2GIS API 3.0: point=lon,lat, radius в метрах, fields=items.point
    params = {
        "q": query_text,
        "point": f"{lon:.6f},{lat:.6f}",
        "radius": max(1000, min(radius_meters, 15000)),
        "page_size": 25,
        "fields": "items.point",
        "key": TWOGIS_API_KEY
    }
    
    found = []
    error_msg = None
    try:
        r = requests.get(url, params=params, timeout=8)
        if r.status_code == 200:
            data = r.json()
            items = data.get("result", {}).get("items", [])
            for it in items:
                point = it.get("point")
                if point and "lat" in point and "lon" in point:
                    found.append({
                        "name": it.get("name", "Объект"),
                        "lat": point["lat"],
                        "lon": point["lon"],
                        "category": cat_name,
                        "address": it.get("address_name", "")
                    })
        else:
            # Запасной запрос через location
            params_fallback = {
                "q": query_text,
                "location": f"{lon:.6f},{lat:.6f}",
                "page_size": 25,
                "fields": "items.point",
                "key": TWOGIS_API_KEY
            }
            r_fb = requests.get(url, params=params_fallback, timeout=8)
            if r_fb.status_code == 200:
                for it in r_fb.json().get("result", {}).get("items", []):
                    pt = it.get("point")
                    if pt and "lat" in pt and "lon" in pt:
                        found.append({
                            "name": it.get("name", "Объект"),
                            "lat": pt["lat"],
                            "lon": pt["lon"],
                            "category": cat_name,
                            "address": it.get("address_name", "")
                        })
            else:
                error_msg = f"HTTP {r.status_code}"
    except Exception as e:
        error_msg = str(e)
        
    return found, error_msg

# 4. Экспорт отчетов
def generate_excel_report(lat, lon, radius_km, score, verdict, attractions, competitors, amenities):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame({
            "Параметр": [
                "Широта (Lat)", "Долгота (Lon)", "Радиус аудита (км)",
                "Итоговый скоринг (из 100)", "Маркетинговый вердикт",
                "Точек притяжения (POI)", "Конкурентов (отели/базы)", "Объектов сервиса (кафе/рестораны)"
            ],
            "Значение": [
                f"{lat:.4f}", f"{lon:.4f}", radius_km,
                score, verdict,
                len(attractions), len(competitors), len(amenities)
            ]
        }).to_excel(writer, sheet_name="Сводка KPI", index=False)

        all_objs = attractions + competitors + amenities
        if all_objs:
            pd.DataFrame(all_objs).to_excel(writer, sheet_name="Реестр объектов 2GIS", index=False)
    return output.getvalue()

def generate_clean_pdf(lat, lon, radius_km, score, verdict_code, poi_count, comp_count, infra_count):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=15)
    pdf.cell(0, 10, "INVESTMENT SITE AUDIT MEMORANDUM (2GIS DATA)", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)

    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 7, f"Coordinates: {lat:.4f}, {lon:.4f} (Republic of Tatarstan)", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 7, f"Buffer Radius: {radius_km} km", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 7, f"Investment Feasibility Score: {score} / 100", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    pdf.set_font("Helvetica", style="B", size=11)
    pdf.cell(0, 7, "Identified Spatial Assets (via 2GIS):", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, f"- Tourist Points of Interest (POI): {poi_count} units", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, f"- Direct Competitors (Hotels/Resorts): {comp_count} units", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, f"- Food & Hospitality Infrastructure: {infra_count} units", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    pdf.set_font("Helvetica", style="B", size=11)
    pdf.cell(0, 7, "Recommendation:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=10)
    pdf.multi_cell(0, 6, verdict_code)
    pdf.ln(8)
    pdf.set_font("Helvetica", style="I", size=8)
    pdf.cell(0, 6, "GeoTour Analytics SPPR Engine. Powered by 2GIS API.", new_x="LMARGIN", new_y="NEXT")
    return bytes(pdf.output())

# 5. Координаты (Казань)
KAZAN_COORDS = [55.7961, 49.1064]

if "target_coords" not in st.session_state:
    st.session_state.target_coords = None

map_center = st.session_state.target_coords if st.session_state.target_coords else KAZAN_COORDS
current_zoom = 13 if st.session_state.target_coords else 11

attractions = []
competitors = []
amenities = []
api_status = ""

if st.session_state.target_coords:
    c_lat, c_lon = st.session_state.target_coords
    with st.spinner("Запрос к официальному каталогу 2GIS Places API..."):
        attractions, err1 = fetch_2gis_places(c_lat, c_lon, radius_m, "музей памятник театр", "Магнит (POI)")
        competitors, err2 = fetch_2gis_places(c_lat, c_lon, radius_m, "гостиница отель", "Конкурент")
        amenities, err3 = fetch_2gis_places(c_lat, c_lon, radius_m, "кафе ресторан", "Инфраструктура")
        
        total_found = len(attractions) + len(competitors) + len(amenities)
        if total_found > 0:
            api_status = f"✅ 2GIS Places API вернул {total_found} реальных объектов в радиусе {radius_km} км"
        elif err1 or err2 or err3:
            api_status = f"⚠️ Статус вызова 2GIS: {err1 or err2 or err3}"
        else:
            api_status = f"ℹ️ В радиусе {radius_km} км объекты не найдены. Увеличьте радиус в левом меню до 5-8 км."

tab_audit, tab_route = st.tabs(["📊 Экспресс-аудит локации", "🗺️ Конструктор экскурсионного маршрута"])

# ==============================================================================
# ВКЛАДКА 1: ЭКСПРЕСС-АУДИТ
# ==============================================================================
with tab_audit:
    if api_status:
        st.info(api_status)

    m_audit = folium.Map(location=map_center, zoom_start=current_zoom, tiles="OpenStreetMap")

    if st.session_state.target_coords:
        lat, lon = st.session_state.target_coords
        folium.Circle(
            location=[lat, lon], radius=radius_m,
            color="#2980b9", weight=2, fill=True, fill_color="#3498db", fill_opacity=0.12
        ).add_to(m_audit)
        folium.Marker([lat, lon], popup="<b>Точка объекта</b>", icon=folium.Icon(color="red", icon="star")).add_to(m_audit)

        cluster_poi = MarkerCluster(name="Магниты (POI)").add_to(m_audit)
        cluster_comp = MarkerCluster(name="Конкуренты").add_to(m_audit)
        cluster_infra = MarkerCluster(name="Инфраструктура").add_to(m_audit)

        if show_poi:
            for p in attractions:
                folium.Marker([p['lat'], p['lon']], popup=f"<b>{p['name']}</b><br>{p.get('address','')}", icon=folium.Icon(color="green", icon="tree-conifer")).add_to(cluster_poi)
        if show_comp:
            for c in competitors:
                folium.Marker([c['lat'], c['lon']], popup=f"<b>{c['name']}</b><br>{c.get('address','')}", icon=folium.Icon(color="orange", icon="bed")).add_to(cluster_comp)
        if show_infra:
            for im in amenities:
                folium.Marker([im['lat'], im['lon']], popup=f"<b>{im['name']}</b><br>{im.get('address','')}", icon=folium.Icon(color="blue", icon="cutlery")).add_to(cluster_infra)

    col_map1, col_rep1 = st.columns([3, 2])
    with col_map1:
        map_out = st_folium(m_audit, width=750, height=540, key="audit_map")
        if map_out and map_out.get("last_clicked"):
            new_pt = [map_out["last_clicked"]["lat"], map_out["last_clicked"]["lng"]]
            if st.session_state.target_coords != new_pt:
                st.session_state.target_coords = new_pt
                st.rerun()

    with col_rep1:
        st.subheader("📊 Аналитический отчет")
        if not st.session_state.target_coords:
            st.info("👈 Нажмите на карту в Казани или окрестностях.")
        else:
            lat, lon = st.session_state.target_coords
            st.write(f"**Координаты площадки:** `{lat:.4f}, {lon:.4f}`")
            
            poi_count = len(attractions)
            comp_count = len(competitors)
            infra_count = len(amenities)

            poi_score = min(poi_count * 5, 45)
            infra_score = min(infra_count * 2, 25)
            transport_score = 15

            if poi_count >= 5:
                market_balance = 15
            elif poi_count >= 1:
                market_balance = 10
            else:
                market_balance = -min(comp_count * 2, 15)

            total_score = max(15, min(100, poi_score + infra_score + transport_score + market_balance))

            m1, m2 = st.columns(2)
            m1.metric("Магниты (POI)", f"{poi_count} шт.")
            m2.metric("Конкуренты", f"{comp_count} шт.")
            m3, m4 = st.columns(2)
            m3.metric("Инфраструктура", f"{infra_count} шт.")
            m4.metric("Скоринг", f"{total_score} / 100")

            st.markdown("---")
            if total_score >= 70:
                if poi_count >= 5:
                    verdict_ru = "ВЫСОКИЙ ПОТЕНЦИАЛ: Плотный городской кластер. Рекомендован отель 3-4* или апарт-комплекс для туристов и деловых гостей."
                    verdict_en = "HIGH POTENTIAL: High-density urban cluster. 3-4 star city hotel or apart-hotel recommended."
                else:
                    verdict_ru = "ВЫСОКИЙ ПОТЕНЦИАЛ: Рекомендован загородный спа-резорт или парк-отель."
                    verdict_en = "HIGH POTENTIAL: Full-scale resort complex recommended."
                st.success(f"🟢 **ВЫСОКИЙ ПОТЕНЦИАЛ**\n\n{verdict_ru}")
            elif total_score >= 45:
                verdict_ru = "УМЕРЕННЫЙ ПОТЕНЦИАЛ: Рекомендован эко-глэмпинг или модульная база отдыха."
                verdict_en = "MODERATE POTENTIAL: Eco-glamping or modular boutique camp recommended."
                st.warning(f"🟡 **УМЕРЕННЫЙ ПОТЕНЦИАЛ**\n\n{verdict_ru}")
            else:
                verdict_ru = "ВЫСОКИЙ РИСК: Дефицит объектов притяжения. Срок окупаемости не прогнозируется."
                verdict_en = "HIGH RISK: Deficiency of tourism magnets."
                st.error(f"🔴 **ВЫСОКИЙ РИСК**\n\n{verdict_ru}")

            st.markdown("---")
            st.subheader("📥 Выгрузка отчетов")
            b_col1, b_col2 = st.columns(2)
            with b_col1:
                excel_bytes = generate_excel_report(lat, lon, radius_km, total_score, verdict_ru, attractions, competitors, amenities)
                st.download_button(
                    "📊 Скачать Excel",
                    data=excel_bytes,
                    file_name=f"site_audit_{lat:.3f}_{lon:.3f}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with b_col2:
                pdf_bytes = generate_clean_pdf(lat, lon, radius_km, total_score, verdict_en, poi_count, comp_count, infra_count)
                st.download_button(
                    "📄 Скачать PDF",
                    data=pdf_bytes,
                    file_name=f"memo_{lat:.3f}_{lon:.3f}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

# ==============================================================================
# ВКЛАДКА 2: КОНСТРУКТОР ТУРА
# ==============================================================================
with tab_route:
    if not st.session_state.target_coords:
        st.info("Сначала выберите базовую точку отеля на карте в первой вкладке.")
    elif len(attractions) < 2:
        st.warning("В данном радиусе мало объектов для составления кругового маршрута. Увеличьте радиус сканирования.")
    else:
        st.subheader("Маршрутизатор кругового экскурсионного тура")
        col_rc, col_rv = st.columns([1, 2])

        with col_rc:
            poi_dict = {f"{p['name']} ({p.get('address','')})": p for p in attractions}
            selected_names = st.multiselect(
                "Объекты для программы тура:",
                options=list(poi_dict.keys()),
                default=list(poi_dict.keys())[:min(4, len(poi_dict))]
            )
            avg_speed = st.slider("Скорость трансфера (км/ч)", 20, 80, 40)
            stop_time_min = st.slider("Время осмотра 1 точки (мин)", 15, 90, 45)

        selected_poi = [poi_dict[name] for name in selected_names]

        if len(selected_poi) >= 1:
            hotel_pt = st.session_state.target_coords
            ordered_poi = solve_tsp_nearest_neighbor(hotel_pt, selected_poi)
            route_coords = [hotel_pt] + [[p['lat'], p['lon']] for p in ordered_poi] + [hotel_pt]

            total_dist_km = sum(
                haversine_distance(route_coords[i][0], route_coords[i][1], route_coords[i+1][0], route_coords[i+1][1]) * 1.3
                for i in range(len(route_coords) - 1)
            )

            drive_hours = total_dist_km / avg_speed
            sightseeing_hours = (len(ordered_poi) * stop_time_min) / 60.0
            total_hours = drive_hours + sightseeing_hours
            balance_ratio = (sightseeing_hours / total_hours * 100) if total_hours > 0 else 0

            with col_rc:
                st.markdown("---")
                st.write(f"**Протяженность кольца:** `{total_dist_km:.1f} км`")
                st.write(f"**Время в пути:** `{drive_hours:.1f} ч`")
                st.write(f"**Время осмотра:** `{sightseeing_hours:.1f} ч`")
                st.write(f"**Всего длительность:** `{total_hours:.1f} ч`")

                if balance_ratio >= 50:
                    st.success(f"✅ Баланс тура: **{balance_ratio:.0f}% впечатлений** (Комфортный тур).")
                else:
                    st.warning(f"⚠️ Баланс тура: **{balance_ratio:.0f}% впечатлений** (Утомительные переезды).")

            with col_rv:
                m_route = folium.Map(location=hotel_pt, zoom_start=current_zoom, tiles="OpenStreetMap")
                folium.PolyLine(
                    locations=route_coords,
                    color="#e74c3c", weight=4, opacity=0.85, dash_array="6",
                    tooltip=f"Маршрут тура ({total_dist_km:.1f} км)"
                ).add_to(m_route)

                folium.Marker(hotel_pt, popup="<b>Базовый отель (Старт/Финиш)</b>", icon=folium.Icon(color="red", icon="home")).add_to(m_route)

                for order_num, p in enumerate(ordered_poi, 1):
                    folium.Marker(
                        [p['lat'], p['lon']],
                        popup=f"<b>Остановка #{order_num}:</b> {p['name']}",
                        tooltip=f"#{order_num}: {p['name']}",
                        icon=folium.Icon(color="green", icon="info-sign")
                    ).add_to(m_route)

                st_folium(m_route, width=750, height=500, key="route_map")

                st.markdown("**Программа тура для клиента:**")
                itinerary_md = f"1. **Выезд:** Базовый отель  \n"
                for step_num, p in enumerate(ordered_poi, 1):
                    itinerary_md += f"{step_num + 1}. **Остановка #{step_num}:** {p['name']} — *{stop_time_min} мин*  \n"
                itinerary_md += f"{len(ordered_poi) + 2}. **Возвращение:** Базовый отель"
                st.info(itinerary_md)
