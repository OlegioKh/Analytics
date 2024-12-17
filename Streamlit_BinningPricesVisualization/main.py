
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import streamlit as st
from io import BytesIO
import numpy as np

url = 'https://docs.google.com/spreadsheets/d/1Ec-ynzQwLDhLiPFDz3JQVc1tN0lSvcN3/export?format=xlsx'

@st.cache_data
def load_data(sheet_url):
    # Load data from the Google Sheet
    data = pd.read_excel(sheet_url, engine='openpyxl')
    return data

data = load_data(url)

# Calculate 'Середня ціна реалізації'
data['Середня ціна реалізації'] = (data['Реалізація, грн.'] / data['Реалізація, к-сть']).round(2)

# Replace infinite and NaN values
data['Середня ціна реалізації'].replace([np.inf, -np.inf], np.nan, inplace=True)
data = data.dropna(subset=['Середня ціна реалізації'])

# Remove rows where 'Реалізація, к-сть' is zero or NaN
data = data[data['Реалізація, к-сть'] > 0]

# Sidebar Filters
st.sidebar.header("Фільтри")
unique_styles = data['Стиль'].dropna().unique()
unique_subgroups = data['ПідГрупа'].dropna().unique()

selected_styles = st.sidebar.multiselect("Оберіть Стиль", options=unique_styles, default=[])
selected_subgroups = st.sidebar.multiselect("Оберіть ПідГрупу", options=unique_subgroups, default=[])

num_bins = st.sidebar.slider("Кількість цінових сегментів", min_value=3, max_value=10, value=4)

# Filtering Data
filtered_data = data.copy()
if selected_styles:
    filtered_data = filtered_data[filtered_data['Стиль'].isin(selected_styles)]
if selected_subgroups:
    filtered_data = filtered_data[filtered_data['ПідГрупа'].isin(selected_subgroups)]

# Drop rows with missing values
filtered_data = filtered_data.dropna(subset=['Реалізація, к-сть', 'Реалізація, грн.', 'Середня ціна реалізації'])

# Ensure only positive 'Реалізація, к-сть' values
filtered_data = filtered_data[filtered_data['Реалізація, к-сть'] > 0]

# Additional NaN check before qcut
if filtered_data['Середня ціна реалізації'].isna().any():
    st.error("Є пропущені значення у 'Середня ціна реалізації'. Перевірте дані.")
else:
    # Price Segments
    try:
        bins = pd.qcut(filtered_data['Середня ціна реалізації'], num_bins, duplicates='drop')
        bin_labels = [f"{interval.left:.2f}-{interval.right:.2f}" for interval in bins.cat.categories]
        filtered_data['Price Segment'] = bins.cat.rename_categories(bin_labels)
    except ValueError as e:
        st.error(f"Помилка при створенні цінових сегментів: {e}")
        filtered_data['Price Segment'] = None

# Aggregated Data for First Chart
aggregated_data = filtered_data.groupby(['ПідГрупа', 'Price Segment']).agg({
    'Реалізація, к-сть': 'sum'
}).reset_index()

# First Chart - Barplot
st.write("### Реалізація по ціновим сегментам")
plt.figure(figsize=(10, 6))
barplot = sns.barplot(
    data=aggregated_data,
    x='ПідГрупа',
    y='Реалізація, к-сть',
    hue='Price Segment'
)
plt.xticks(rotation=45)
plt.title("Реалізація к-сть по цінових сегментах")
plt.ylabel("Реалізація, к-сть")

# Annotate bar values
for p in barplot.patches:
    barplot.annotate(f'{int(p.get_height())}',
                     (p.get_x() + p.get_width() / 2., p.get_height()),
                     ha='center', va='center', xytext=(0, 10),
                     textcoords='offset points')
st.pyplot(plt)

# Second Chart - Pie Chart
st.write("### Питома вага цінових сегментів в загальному доході")
pie_data = filtered_data.groupby('Price Segment')['Дохід, грн.'].sum()
if not pie_data.empty:
    plt.figure(figsize=(8, 8))
    pie_labels = [f"{segment}\n{value:,.0f} грн" for segment, value in zip(pie_data.index, pie_data.values)]
    plt.pie(pie_data, labels=pie_labels, autopct='%1.1f%%', startangle=140)
    plt.title("Питома вага цінових сегментів в доході")
    st.pyplot(plt)
else:
    st.write("Немає даних для побудови графіка.")

# Table with Results
st.write("### Деталізована таблиця")
required_columns = ['Артикул - назва', 'Реалізація, к-сть', 'Середня ціна реалізації', 'Price Segment', 'Постачальник', 'Дохід, грн.']

if all(col in filtered_data.columns for col in required_columns):
    # Групування та агрегація даних
    grouped_data = filtered_data.groupby(['Price Segment', 'Постачальник', 'Артикул - назва'], observed=False).agg({
        'Реалізація, к-сть': 'sum',
        'Реалізація, грн.': 'sum',
        'Дохід, грн.': 'sum',
        'Середня ціна реалізації': 'mean'
    }).reset_index()

    # Видаляємо рядки з нульовими значеннями після агрегації
    grouped_data = grouped_data[grouped_data['Реалізація, к-сть'] > 0]

    # Перейменування колонок
    grouped_data = grouped_data.rename(columns={
        'Price Segment': 'Сегмент ціни',
        'Постачальник': 'Постачальник',
        'Артикул - назва': 'Товар',
        'Реалізація, к-сть': 'Реалізація кіл-сть',
        'Реалізація, грн.': 'Реалізація ЦР',
        'Дохід, грн.': 'Дохід',
        'Середня ціна реалізації': 'Середня ціна'
    })

    # Сортування за доходом та сегментом ціни
    grouped_data = grouped_data.sort_values(
        by=['Сегмент ціни', 'Дохід'],
        ascending=[False, False]
    )

    # Форматування даних
    grouped_data['Реалізація ЦР'] = grouped_data['Реалізація ЦР'].astype(int)
    grouped_data['Реалізація кіл-сть'] = grouped_data['Реалізація кіл-сть'].astype(int)
    grouped_data['Дохід'] = grouped_data['Дохід'].astype(int)
    grouped_data['Середня ціна'] = grouped_data['Середня ціна'].round(2)

    # Відображення таблиці
    st.dataframe(
        grouped_data.head(100).style.format({
            'Реалізація ЦР': "{:,.0f}",
            'Реалізація кіл-сть': "{:,.0f}",
            'Дохід': "{:,.0f}",
            'Середня ціна': "{:.2f}"
        }),
        use_container_width=True,
        height=500
    )

    # Генерація Excel
    @st.cache_data
    def to_excel(dataframe):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            dataframe.to_excel(writer, index=False, sheet_name='Sheet1')
        return output.getvalue()

    download_data = to_excel(grouped_data)
    st.download_button(
        label="Завантажити таблицю в Excel",
        data=download_data,
        file_name='detailed_table.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.error("Відсутні необхідні колонки для побудови таблиці.")


