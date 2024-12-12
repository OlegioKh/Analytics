import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import streamlit as st
from io import BytesIO

st.markdown(
    """
    <style>
    .main {
        max-width: 1200px;
        margin: 0 auto;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

def to_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

@st.cache_data
def load_data(file_path):
    return pd.read_excel(file_path)

file_path = 'DB - SeaBorn.xlsx'
data = pd.read_excel(file_path)

st.sidebar.header("Filters")
styles = data['Стиль'].unique()
subgroups = data['ПідГрупа'].unique()

selected_styles = st.sidebar.multiselect("Select Style(s)", options=styles, default=[])
selected_subgroups = st.sidebar.multiselect("Select Subgroup(s)", options=subgroups, default=[])

num_bins = st.sidebar.slider("Number of Price Bins", min_value=3, max_value=10, value=4)

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

filtered_data = filtered_data.dropna(subset=['Середня ЦР'])

bins = pd.qcut(
    filtered_data['Середня ЦР'],
    num_bins,
    duplicates='drop'
)

bin_ranges = bins.cat.categories
new_labels = [f"{int(interval.left)}-{int(interval.right)}" for interval in bin_ranges]
filtered_data['Price Segment'] = bins.cat.rename_categories(new_labels)

# Визначаємо, що використовувати на осі X
if selected_styles and not selected_subgroups:
    group_by_column = 'Стиль'
elif selected_subgroups and not selected_styles:
    group_by_column = 'ПідГрупа'
else:
    group_by_column = 'Стиль'  # За замовчуванням використовуємо стиль

aggregated_data = filtered_data.groupby([group_by_column, 'Price Segment']).agg({
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
    plt.xlabel("Style", fontsize=14)
    plt.ylabel("Sales Quantity", fontsize=14)
    plt.title("Sales Quantity by Style and Price Segment", fontsize=18)
    st.pyplot(plt)

    # Pie Chart
    pie_data = aggregated_data.groupby('Price Segment')['Дохід, грн.'].sum()
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
        ax.set_title("Income Distribution by Price Segment", fontsize=16)
        st.pyplot(fig)
    else:
        st.write("No data available for pie chart.")

# Таблиця
with col2:
    if all(col in filtered_data.columns for col in ['Артикул - назва', 'Реалізація, к-сть', 'Середня ЦР', 'Price Segment', 'Постачальник']):
        # Обчислюємо загальні продажі для кожного цінового сегмента
        segment_totals = filtered_data.groupby('Price Segment')['Реалізація, к-сть'].sum().reset_index()
        segment_totals = segment_totals.rename(columns={'Реалізація, к-сть': 'Total Sales'})

        # Додаємо загальні продажі сегмента до основних даних
        filtered_data = pd.merge(filtered_data, segment_totals, on='Price Segment', how='left')

        # Pivot table з сортуванням
        pivot_table = filtered_data.pivot_table(
            index=['Price Segment', 'Постачальник', 'Артикул - назва'],
            values=['Реалізація, к-сть', 'Реалізація, грн.', 'Середня ЦР', 'Total Sales'],
            aggfunc={
                'Реалізація, к-сть': 'sum',
                'Реалізація, грн.': 'sum',  # Додано 'Реалізація, грн.'
                'Середня ЦР': 'mean',
                'Total Sales': 'max'
            }
        ).reset_index()

        # Сортуємо спочатку за загальними продажами сегмента, потім за кількістю реалізації товару
        pivot_table = pivot_table.sort_values(by=['Total Sales', 'Реалізація, к-сть'], ascending=False)

        # Видаляємо колонку Total Sales з таблиці
        pivot_table = pivot_table.drop(columns=['Total Sales'])

        # Перейменовуємо колонки для відображення
        pivot_table = pivot_table.rename(columns={
            'Price Segment': 'Price Segment',
            'Постачальник': 'Supplier',
            'Артикул - назва': 'Item Name',
            'Реалізація, к-сть': 'Sales Quantity',
            'Реалізація, грн.': 'Sales Amount',
            'Середня ЦР': 'Average Price'
        })

        # Форматуємо колонки
        pivot_table['Sales Amount'] = pivot_table['Sales Amount'].astype(int)  # Перетворення у формат int
        pivot_table['Average Price'] = pivot_table['Average Price'].round(2)  # Округлення до 2 знаків після коми

        if len(pivot_table) > 1000:  # Ліміт на 200 рядків
            st.write("The table is too large to display. Showing the first 200 rows.")
            pivot_table = pivot_table.head(200)

        # Відображаємо таблицю в Streamlit
        st.dataframe(
            pivot_table.style
            .format({'Average Price': "{:.2f}", 'Sales Amount': "{:.0f}"})  # Формат для числових значень
            .set_properties(subset=['Supplier', 'Item Name'], width="50%")
            .set_properties(subset=['Sales Quantity', 'Sales Amount', 'Average Price'], width="25%"),
            use_container_width=True,
            height=500
        )

        excel_data = to_excel(pivot_table)
        st.download_button(
            label="Download full table as Excel",
            data=excel_data,
            file_name='pivot_table.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    else:
        st.write("Columns 'Артикул - назва', 'Реалізація, к-сть', 'Середня ЦР', 'Price Segment', or 'Постачальник' are missing in the dataset.")

