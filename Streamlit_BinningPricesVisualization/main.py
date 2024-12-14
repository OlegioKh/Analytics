import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import streamlit as st
from io import BytesIO


url = 'https://docs.google.com/spreadsheets/d/1Ec-ynzQwLDhLiPFDz3JQVc1tN0lSvcN3/export?format=xlsx'

@st.cache_data
def load_data(sheet_url):
    # Завантажуємо файл напряму
    data = pd.read_excel(sheet_url, engine='openpyxl')
    return data

# Виклик функції для завантаження даних
data = load_data(url)

st.write("**Дані з Google Sheets**:")
sorted_data = data.sort_values(by="Реалізація, грн.", ascending=False)
st.write(f"Кількість рядків у датасеті: {len(sorted_data)}")

for col in ['Реалізація, к-сть', 'Реалізація, грн.', 'Дохід, грн.', 'Середня ЦР']:
    if col in data.columns:
        data[col] = pd.to_numeric(data[col], errors='coerce')


def to_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

excel_data = to_excel(data)

st.sidebar.header("Filters")
styles = data['Стиль'].unique()
subgroups = data['ПідГрупа'].unique()

if len(styles) > 0:  # Перевірка, чи є доступні стилі
    selected_styles = st.sidebar.multiselect("Select Style(s)", options=styles, default=[])
else:
    st.sidebar.write("Немає доступних стилів для вибору.")

if len(subgroups) > 0:  # Перевірка, чи є доступні підгрупи
    selected_subgroups = st.sidebar.multiselect("Select Subgroup(s)", options=subgroups, default=[])
else:
    st.sidebar.write("Немає доступних підгруп для вибору.")

num_bins = st.sidebar.slider("Number of Price Bins", min_value=3, max_value=6, value=4)

# Фільтрація даних на основі вибраних стилів і підгруп
if selected_styles and selected_subgroups:
    filtered_data = data[
        (data['Стиль'].isin(selected_styles)) & (data['ПідГрупа'].isin(selected_subgroups))
    ]
elif selected_styles:
    filtered_data = data[data['Стиль'].isin(selected_styles)]
elif selected_subgroups:
    filtered_data = data[data['ПідГрупа'].isin(selected_subgroups)]
else:
    filtered_data = data


filtered_data = filtered_data.dropna(subset=['Реалізація, к-сть', 'Реалізація, грн.', 'Дохід, грн.', 'Середня ЦР'])

bins = pd.qcut(
    filtered_data['Середня ЦР'],
    num_bins,
    duplicates='drop'
)

bin_ranges = bins.cat.categories
new_labels = [f"{int(interval.left)}-{int(interval.right)}" for interval in bin_ranges]

#filtered_data = filtered_data.copy()  # Робимо повну копію DataFrame
filtered_data['Price Segment'] = bins.cat.rename_categories(new_labels)

# Визначаємо, що використовувати на осі X
if selected_styles and not selected_subgroups:
    group_by_column = 'Стиль'
elif selected_subgroups and not selected_styles:
    group_by_column = 'ПідГрупа'
else:
    group_by_column = 'Стиль'  # За замовчуванням використовуємо стиль

aggregated_data = filtered_data.groupby([group_by_column, 'Price Segment'], observed=False).agg({
    'Реалізація, к-сть': 'sum',
    'Реалізація, грн.': 'sum',
    'Дохід, грн.': 'sum'
}).reset_index()

col1, col2 = st.columns(2)

# Графік
with col1:
    plt.figure(figsize=(11, 6))
    ax = sns.barplot(
        data=aggregated_data,
        x=group_by_column,
        y='Реалізація, к-сть',
        hue='Price Segment'
    )

    legend = ax.legend(title='Price Segment', fontsize=14, title_fontsize=16)
    legend.set_bbox_to_anchor((1.05, 1))

    for container in ax.containers:
        ax.bar_label(container, fmt='%d', label_type='edge', fontsize=14)
    plt.xticks(rotation=45, fontsize=14)
    plt.yticks(fontsize=14)
    plt.xlabel("Структура", fontsize=14)
    plt.ylabel("Реалізація кіл-сть", fontsize=14)
    plt.title("Реалізація в кількості залежно від цінового сегменту", fontsize=18)
    st.pyplot(plt)

    # Pie Chart
    pie_data = aggregated_data.groupby('Price Segment', observed=False)['Дохід, грн.'].sum()
    pie_data = pie_data[pie_data > 0]
    if not pie_data.empty:
        total_sales = pie_data.sum()
        pie_labels = [f"{segment}\n{value:,.0f} грн ({value / total_sales * 100:.1f}%)"
                      for segment, value in pie_data.items()]  # Додано суму реалізації

        fig, ax = plt.subplots(figsize=(8, 8))
        ax.pie(
            pie_data,
            labels=pie_labels,
            autopct=None,  # Відключає автоматичний відсоток
            startangle=140,
            colors=sns.color_palette('pastel'),
        )
        ax.set_title("Дохід залежно від цінового сегменту", fontsize=16)
        st.pyplot(fig)
    else:
        st.write("No data available for pie chart.")

# Таблиця
with col2:
    required_columns = ['Артикул - назва', 'Реалізація, к-сть', 'Середня ЦР', 'Price Segment', 'Постачальник']

    # Перевірка наявності необхідних колонок
    if all(col in filtered_data.columns for col in required_columns):
        # Групування даних
        grouped_data = filtered_data.groupby(['Price Segment', 'Постачальник', 'Артикул - назва'], observed=False).agg({
            'Реалізація, к-сть': 'sum',
            'Реалізація, грн.': 'sum',
            'Середня ЦР': 'mean'
        }).reset_index()

        # Перейменовуємо колонки для відображення
        grouped_data = grouped_data.rename(columns={
            'Price Segment': 'Сегмент ціни',
            'Постачальник': 'Постачальник',
            'Артикул - назва': 'Товар',
            'Реалізація, к-сть': 'Реалізація кіл-сть',
            'Реалізація, грн.': 'Реалізація ЦР',
            'Середня ЦР': 'Середня ціна'
        })

        # Сортуємо дані
        grouped_data = grouped_data.sort_values(
            by=['Сегмент ціни', 'Реалізація ЦР', 'Товар'],  # Сортування за декількома стовпцями
            ascending=[True, False, True]  # 'Сегмент ціни' - зростання, 'Реалізація ЦР' - спадання, 'Товар' - зростання
        )

        # Форматування даних
        grouped_data['Реалізація ЦР'] = grouped_data['Реалізація ЦР'].astype(int)
        grouped_data['Реалізація кіл-сть'] = grouped_data['Реалізація кіл-сть'].astype(int)
        grouped_data['Середня ціна'] = grouped_data['Середня ціна'].round(2)

        # Відображення таблиці
        st.dataframe(
            grouped_data.head(100).style
            .format({
                'Реалізація ЦР': "{:,.0f}",
                'Реалізація кіл-сть': "{:,.0f}",
                'Середня ціна': "{:.2f}"
            }),
            use_container_width=True,
            height=500
        )

        # Генерація Excel для завантаження
        excel_data = to_excel(grouped_data)
        st.download_button(
            label="Завантажити таблицю в Excel",
            data=excel_data,
            file_name='table.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    else:
        st.error("Відсутні необхідні колонки для побудови таблиці.")

