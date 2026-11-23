from datetime import date
from io import BytesIO
from math import isfinite
import random

import pandas as pd
import plotly.express as px
import streamlit as st


DATA_COLUMNS = ["Дата", "Кінотеатр", "Жанр", "Відвідувачі", "Сеанси", "Дохід, грн"]
FILTER_KEYS = ("filter_cinemas", "filter_genres", "filter_dates")


st.set_page_config(
    page_title="Аналіз кінотеатру",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)


@st.cache_data
def create_cinema_data() -> pd.DataFrame:
    """Створює демонстраційні дані про відвідуваність кінотеатрів."""
    rng = random.Random(42)
    cinemas = ["Cinema City", "Movie Park", "Star Hall"]
    genres = ["Комедія", "Пригоди", "Фантастика", "Драма", "Анімація"]
    base_price = {
        "Комедія": 180,
        "Пригоди": 210,
        "Фантастика": 230,
        "Драма": 170,
        "Анімація": 160,
    }
    rows = []

    for day in pd.date_range(end=date.today(), periods=180):
        weekend_bonus = 45 if day.dayofweek >= 5 else 0
        for cinema in cinemas:
            cinema_bonus = cinemas.index(cinema) * 14
            for genre in genres:
                visitors = rng.randint(35, 145) + weekend_bonus + cinema_bonus
                sessions = rng.randint(2, 6)
                price = base_price[genre]
                rows.append(
                    {
                        "Дата": day,
                        "Кінотеатр": cinema,
                        "Жанр": genre,
                        "Відвідувачі": visitors,
                        "Сеанси": sessions,
                        "Дохід, грн": visitors * price,
                    }
                )

    return pd.DataFrame(rows)


def hryvnias(value: float) -> str:
    return f"{value:,.0f}".replace(",", " ") + " грн"


def excel_template() -> bytes:
    example = pd.DataFrame(
        [(date(2026, 1, 15), "Cinema City", "Комедія", 95, 3, 17100)],
        columns=DATA_COLUMNS,
    )
    buffer = BytesIO()
    example.to_excel(buffer, engine="openpyxl", index=False, sheet_name="Відвідуваність")
    return buffer.getvalue()


def validate_excel_data(raw: pd.DataFrame) -> pd.DataFrame:
    raw.columns = raw.columns.str.strip()
    missing = [column for column in DATA_COLUMNS if column not in raw.columns]
    if missing:
        raise ValueError("У таблиці бракує стовпців — " + ", ".join(missing) + ".")
    duplicates = raw.columns[raw.columns.duplicated()].intersection(DATA_COLUMNS)
    if not duplicates.empty:
        raise ValueError("У таблиці повторюються заголовки — " + ", ".join(duplicates) + ".")

    data = raw[DATA_COLUMNS].replace(r"^\s*$", pd.NA, regex=True).dropna(how="all").copy()
    if data.empty:
        raise ValueError("Додайте хоча б один запис під заголовками таблиці.")

    # Excel також може зберігати дати як текст.
    dates = pd.to_datetime(
        data["Дата"].astype("string").str.strip(),
        format="mixed",
        dayfirst=True,
        errors="coerce",
        utc=True,
    )
    invalid = {"Дата": dates.isna()}
    data["Дата"] = dates.dt.tz_localize(None).dt.normalize()

    for column in ("Кінотеатр", "Жанр"):
        data[column] = data[column].astype("string").str.strip()
        invalid[column] = data[column].isna()

    for column in ("Відвідувачі", "Сеанси", "Дохід, грн"):
        cells = data[column].astype("string").str.strip().str.replace(",", ".", regex=False)
        values = pd.to_numeric(cells, errors="coerce")
        valid = values.ge(0) & values.lt(float("inf"))
        if column != "Дохід, грн":
            valid &= values.mod(1).eq(0) & values.le(2**53 - 1)
        invalid[column] = ~valid.fillna(False)
        data[column] = values

    errors = []
    for column, mask in invalid.items():
        if mask.any():
            rows = [str(index + 2) for index in data.index[mask]]
            suffix = " та інші" if len(rows) > 5 else ""
            errors.append(f"«{column}» (рядки {', '.join(rows[:5])}{suffix})")
    if errors:
        raise ValueError("Перевірте значення у стовпцях — " + ", ".join(errors) + ".")

    return data.astype({"Відвідувачі": "int64", "Сеанси": "int64", "Дохід, грн": "float64"})


def reset_filters() -> None:
    for key in FILTER_KEYS:
        st.session_state.pop(key, None)


def manual_data() -> pd.DataFrame:
    with st.form("cinema_entry", clear_on_submit=False):
        first, second, third = st.columns(3)
        entry_date = first.date_input("Дата", value=date.today(), format="DD.MM.YYYY")
        cinema = second.text_input("Кінотеатр").strip()
        genre = third.text_input("Жанр").strip()
        visitors = first.number_input("Відвідувачі", min_value=0, step=1)
        sessions = second.number_input("Сеанси", min_value=0, step=1)
        revenue = third.number_input("Дохід, грн", min_value=0.0, step=0.01, format="%.2f")
        submitted = st.form_submit_button("Додати запис")

    records = st.session_state.setdefault("cinema_records", [])
    if submitted:
        if not cinema or not genre or not isfinite(revenue):
            st.error("Заповніть назву кінотеатру й жанр та введіть коректну суму доходу.")
        else:
            records.append((pd.Timestamp(entry_date), cinema, genre, visitors, sessions, revenue))
            reset_filters()
            st.success(f"Запис додано. Усього записів — {len(records)}.")

    st.caption("Введені записи зберігаються протягом поточного сеансу роботи.")
    return pd.DataFrame(records, columns=DATA_COLUMNS)


def uploaded_data() -> pd.DataFrame:
    st.caption(
        "Програма читає перший аркуш .xlsx. Скопіюйте заголовки із шаблону. "
        "Дату можна ввести як дату Excel або текст у форматі ДД.ММ.РРРР. "
        "Відвідувачі й сеанси мають бути цілими невід’ємними числами."
    )
    st.download_button(
        "Завантажити шаблон Excel",
        data=excel_template(),
        file_name="cinema_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    uploaded = st.file_uploader("Виберіть файл Excel", type="xlsx", on_change=reset_filters)
    if uploaded is None:
        st.stop()

    try:
        raw = pd.read_excel(uploaded, engine="openpyxl", keep_default_na=False)
    except Exception:
        st.error("Не вдалося прочитати файл. Перевірте, чи він відкривається в Excel як .xlsx.")
        st.stop()

    try:
        return validate_excel_data(raw)
    except ValueError as error:
        st.error(str(error))
        st.stop()


st.markdown(
    """
    <style>
        .stApp { background: #11100f; color: #f6f4ef; }
        [data-testid="stHeader"] { background: transparent; }
        [data-testid="stSidebar"] { background: #1d1b18; }
        [data-testid="stSidebar"] * { color: #f6f4ef; }
        .block-container { padding-top: 3rem; padding-bottom: 3rem; }
        h1, h2, h3, p, label { color: #f6f4ef !important; }
        h1 { font-weight: 800 !important; letter-spacing: -0.04em; }
        [data-testid="stMetric"] {
            background: #25231f;
            border: 1px solid #4b4840;
            border-radius: 0.75rem;
            padding: 1rem;
        }
        [data-testid="stMetricLabel"] p,
        [data-testid="stMetricValue"] { color: #f6f4ef !important; }
        [data-baseweb="select"] > div,
        [data-testid="stDateInput"] input {
            background: #2a2824 !important;
            border-color: #4b4840 !important;
            color: #f6f4ef !important;
        }
        .stDownloadButton > button {
            background: #d44c5b;
            border: 0;
            color: white;
            font-weight: 600;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Аналіз відвідуваності кінотеатрів 🎬")
source = st.sidebar.radio(
    "Джерело даних",
    ("Демонстраційні дані", "Ввести вручну", "Завантажити Excel"),
    on_change=reset_filters,
)

if source == "Ввести вручну":
    data = manual_data()
    if data.empty:
        st.info("Додайте перший запис, щоб побачити статистику.")
        st.stop()
elif source == "Завантажити Excel":
    data = uploaded_data()
else:
    st.caption("Демонстраційні дані за останні 180 днів.")
    data = create_cinema_data()

with st.sidebar:
    st.header("Фільтри")
    selected_cinemas = st.multiselect(
        "Кінотеатри",
        options=sorted(data["Кінотеатр"].unique()),
        default=sorted(data["Кінотеатр"].unique()),
        key=FILTER_KEYS[0],
    )
    selected_genres = st.multiselect(
        "Жанри",
        options=sorted(data["Жанр"].unique()),
        default=sorted(data["Жанр"].unique()),
        key=FILTER_KEYS[1],
    )
    dates = st.date_input(
        "Період",
        value=(data["Дата"].min().date(), data["Дата"].max().date()),
        min_value=data["Дата"].min().date(),
        max_value=data["Дата"].max().date(),
        key=FILTER_KEYS[2],
    )

if len(dates) != 2:
    st.info("Оберіть початкову та кінцеву дату для перегляду статистики.")
    st.stop()

filtered = data[
    data["Кінотеатр"].isin(selected_cinemas)
    & data["Жанр"].isin(selected_genres)
    & data["Дата"].between(pd.Timestamp(dates[0]), pd.Timestamp(dates[1]))
].copy()

if filtered.empty:
    st.warning("За вибраними фільтрами даних немає. Змініть параметри ліворуч.")
    st.stop()

total_visitors = filtered["Відвідувачі"].sum()
total_revenue = filtered["Дохід, грн"].sum()
total_sessions = filtered["Сеанси"].sum()
top_genre = filtered.groupby("Жанр")["Відвідувачі"].sum().idxmax()

metric_1, metric_2, metric_3, metric_4 = st.columns(4)
metric_1.metric("Відвідувачі", f"{total_visitors:,}".replace(",", " "))
metric_2.metric("Дохід", hryvnias(total_revenue))
metric_3.metric("Сеанси", f"{total_sessions:,}".replace(",", " "))
metric_4.metric("Популярний жанр", top_genre)

daily_visitors = filtered.groupby("Дата", as_index=False)["Відвідувачі"].sum()
genre_summary = (
    filtered.groupby("Жанр", as_index=False)[["Відвідувачі", "Дохід, грн"]].sum().sort_values("Відвідувачі")
)
cinema_revenue = filtered.groupby("Кінотеатр", as_index=False)["Дохід, грн"].sum()

left, right = st.columns((1.55, 1))
with left:
    st.subheader("Відвідуваність за днями")
    daily_chart = px.line(
        daily_visitors,
        x="Дата",
        y="Відвідувачі",
        markers=len(daily_visitors) == 1,
        template="plotly_dark",
        color_discrete_sequence=["#e95a6a"],
    )
    daily_chart.update_layout(
        height=330,
        margin=dict(l=0, r=0, t=10, b=0),
        paper_bgcolor="#11100f",
        plot_bgcolor="#11100f",
    )
    st.plotly_chart(daily_chart, use_container_width=True)

with right:
    st.subheader("Дохід за кінотеатрами")
    cinema_chart = px.pie(
        cinema_revenue,
        names="Кінотеатр",
        values="Дохід, грн",
        hole=0.58,
        template="plotly_dark",
        color_discrete_sequence=["#e95a6a", "#f0b86b", "#76a5df"],
    )
    cinema_chart.update_layout(
        height=330,
        margin=dict(l=0, r=0, t=10, b=0),
        paper_bgcolor="#11100f",
        legend=dict(orientation="h", y=-0.15),
    )
    st.plotly_chart(cinema_chart, use_container_width=True)

st.subheader("Популярність жанрів")
genre_chart = px.bar(
    genre_summary,
    x="Відвідувачі",
    y="Жанр",
    orientation="h",
    text="Відвідувачі",
    template="plotly_dark",
    color="Відвідувачі",
    color_continuous_scale=["#7b2630", "#e95a6a", "#f0b86b"],
)
genre_chart.update_layout(
    height=310,
    margin=dict(l=0, r=0, t=10, b=0),
    paper_bgcolor="#11100f",
    plot_bgcolor="#11100f",
    coloraxis_showscale=False,
)
st.plotly_chart(genre_chart, use_container_width=True)

with st.expander("Переглянути дані"):
    st.dataframe(
        filtered.sort_values("Дата", ascending=False),
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        "Завантажити відфільтровані дані у CSV",
        data=filtered.to_csv(index=False).encode("utf-8-sig"),
        file_name="відвідуваність_кінотеатрів.csv",
        mime="text/csv",
    )
