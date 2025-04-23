import streamlit as st
import pandas as pd
import io

st.title("Обновлённая адресная программа СДЭК с округами и изменениями")

st.markdown("""
Загрузите два Excel-файла:
- **Старый файл** — содержит вручную заполненный столбец `Округ` на листе `СДЭК`
- **Новый файл** — свежая адресная программа от СДЭК
""")

old_file = st.file_uploader("Старый файл (например, 07.04)", type="xlsx")
new_file = st.file_uploader("Новый файл (например, 09.04)", type="xlsx")

if old_file:
    xls = pd.ExcelFile(old_file)
    cdek_sheets = [name for name in xls.sheet_names if name.strip().lower().startswith("сдэк")]

    if not cdek_sheets:
        st.error("В старом файле не найдено ни одного листа, начинающегося с 'СДЭК'.")
        st.stop()
    elif len(cdek_sheets) == 1:
        selected_sheet = cdek_sheets[0]
        st.info(f"Автоматически выбран лист: {selected_sheet}")
    else:
        selected_sheet = st.selectbox("Выберите лист из старого файла:", cdek_sheets)

if old_file and new_file:
    try:
        df_old = pd.read_excel(xls, sheet_name=selected_sheet, skiprows=32)
        df_new = pd.read_excel(new_file, skiprows=32)

        required_cols = {'GID', 'Адрес', 'Средняя проходимость месяц'}
        if not required_cols.issubset(df_old.columns) or not required_cols.issubset(df_new.columns):
            st.error("Оба файла должны содержать колонки: GID, Адрес, Средняя проходимость месяц")
        else:
            gid_to_district = df_old.set_index('GID')['Округ'].to_dict()
            insert_index = df_new.columns.get_loc('GID') + 1
            df_new.insert(insert_index, 'Округ', df_new['GID'].map(gid_to_district))

            st.subheader("Обновлённая таблица с округами")
            st.dataframe(df_new)

            buffer_full = io.BytesIO()
            df_new.to_excel(buffer_full, index=False, engine='openpyxl')
            st.download_button(
                label="📥 Скачать полную таблицу с округами",
                data=buffer_full.getvalue(),
                file_name="адресная_программа_с_округами.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Сравнение изменений
            df_old_renamed = df_old[['GID', 'Адрес', 'Средняя проходимость месяц']].copy()
            df_new_renamed = df_new[['GID', 'Адрес', 'Средняя проходимость месяц']].copy()

            merged = df_old_renamed.merge(
                df_new_renamed,
                on='GID',
                how='outer',
                suffixes=('_старый', '_новый'),
                indicator=True
            )

            def detect_change(row):
                if row['_merge'] == 'left_only':
                    return 'Удалено'
                elif row['_merge'] == 'right_only':
                    return 'Добавлено'
                elif (row['Адрес_старый'] != row['Адрес_новый'] or
                      row['Средняя проходимость месяц_старый'] != row['Средняя проходимость месяц_новый']):
                    return 'Изменено'
                else:
                    return 'Без изменений'

            merged['Тип изменения'] = merged.apply(detect_change, axis=1)
            diff_df = merged[merged['Тип изменения'] != 'Без изменений']

            st.subheader("Изменения по сравнению со старым файлом")
            st.dataframe(diff_df)

            buffer_diff = io.BytesIO()
            diff_df.to_excel(buffer_diff, index=False, engine='openpyxl')
            st.download_button(
                label="📥 Скачать таблицу изменений",
                data=buffer_diff.getvalue(),
                file_name="изменения_по_сравнению.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ошибка обработки: {e}")
