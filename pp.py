import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(
    page_title="Analitički Alat za Izvještaje",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Analitički Alat za Izvještaje o Rezervacijama")
st.markdown("Učitajte vaš Excel izvještaj kako biste dobili detaljnu analizu i vizualizaciju ključnih poslovnih metrika.")

def process_data(df):
    price_cols = ['Net Price', 'Sale Price', 'Agency Payment', 'Passenger Amount to Pay', 'Agency Amount to Pay', 'Profit']
    for col in price_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    date_cols = ['Create Date', 'Begin Date', 'End Date']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors='coerce')

    df['Package Type'].fillna('Grupni polazak', inplace=True)
    df['Package Type'] = df['Package Type'].replace({'individual': 'Individualni polazak'})

    df.fillna({
        'Net Price': 0, 'Sale Price': 0, 'Agency Payment': 0, 'Passenger Amount to Pay': 0, 
        'Agency Amount to Pay': 0, 'Profit': 0, 'Night': 0, 'Adult': 0, 'Child': 0, 'Infant': 0
    }, inplace=True)
    
    df['Profit'] = df['Agency Amount to Pay'] - df['Net Price']
    
    df['Total Pax'] = df['Adult'] + df['Child'] + df['Infant']
    
    return df

def to_excel(df_filtered, kpi_summary):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_filtered.to_excel(writer, index=False, sheet_name='Analizirani Podaci')
        kpi_df = pd.DataFrame(kpi_summary.items(), columns=['Metrika', 'Vrijednost'])
        kpi_df.to_excel(writer, index=False, sheet_name='Ključne Metrike (KPI)')
        if not df_filtered.empty:
            profit_by_city = df_filtered.groupby('Arrival City')['Profit'].sum().sort_values(ascending=False).reset_index()
            profit_by_city.to_excel(writer, index=False, sheet_name='Profit po Destinaciji')
            top_hotels = df_filtered['Hotel Name'].value_counts().reset_index()
            top_hotels.columns = ['Hotel', 'Broj Rezervacija']
            top_hotels.to_excel(writer, index=False, sheet_name='Najprodavaniji Hoteli')
            profit_by_package = df_filtered.groupby('Package Type')['Profit'].sum().reset_index()
            profit_by_package.to_excel(writer, index=False, sheet_name='Profit po Tipu Paketa')
            profit_by_author = df_filtered.groupby('Author')['Profit'].sum().sort_values(ascending=False).reset_index()
            profit_by_author.to_excel(writer, index=False, sheet_name='Profit po Autoru')
    processed_data = output.getvalue()
    return processed_data

uploaded_file = st.file_uploader("Odaberite Excel dokument (.xlsx ili .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        required_columns = [
            "Reservation No", "Arrival City", "Hotel Name", "Author", "Payment", 
            "Agency", "Begin Date", "Package", "Price List", "End Date", "Night", "Adult", 
            "Child", "Infant", "Net Price", "Sale Price", "Agency Payment", 
            "Create Date", "Passenger Amount to Pay", "Agency Amount to Pay", 
            "Package Type", "Profit"
        ]
        
        cleaned_required_columns = [col.strip() for col in required_columns]

        def find_header_row(file, required_cols, max_rows=10):
            for i in range(max_rows):
                file.seek(0)
                try:
                    temp_df = pd.read_excel(file, nrows=1, header=None, skiprows=i)
                    if not temp_df.empty and not temp_df.iloc[0].isnull().all():
                        header_row = temp_df.iloc[0].astype(str).str.strip().tolist()
                        if all(col.strip() in header_row for col in required_cols):
                            return i
                except Exception:
                    continue
            return None

        uploaded_file.seek(0)
        
        header_row_index = find_header_row(uploaded_file, required_columns)

        if header_row_index is not None:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=header_row_index)
            st.success(f"Fajl uspješno učitan. Zaglavlje pronađeno u redu: {header_row_index + 1}")

            df.columns = df.columns.str.strip()
            
            missing_columns = [col for col in cleaned_required_columns if col not in df.columns]

            if missing_columns:
                st.error(f"Greška: Nakon učitavanja, nedostaju sljedeće kolone:")
                for col in missing_columns:
                    st.error(f"- **{col}**")
                st.warning("Molimo provjerite da li su nazivi kolona u Excelu potpuno identični s onima na listi.")
            else:
                st.success("Sve potrebne kolone su pronađene!")
                
                if 'df_state' not in st.session_state:
                    df = process_data(df)
                    df['Uključi u analizu'] = (df['Net Price'] > 0) & (df['Agency Amount to Pay'] > 0)
                    st.session_state.df_state = df.copy()

                st.sidebar.header("Filteri za Analizu")
                
                all_cities = st.session_state.df_state['Arrival City'].dropna().unique()
                selected_cities = st.sidebar.multiselect("Destinacija (Arrival City)", options=sorted(all_cities), default=sorted(all_cities))
                all_hotels = st.session_state.df_state['Hotel Name'].dropna().unique()
                selected_hotels = st.sidebar.multiselect("Hotel", options=sorted(all_hotels), default=sorted(all_hotels))
                all_authors = st.session_state.df_state['Author'].dropna().unique()
                selected_authors = st.sidebar.multiselect("Autor", options=sorted(all_authors), default=sorted(all_authors))
                all_agencies = st.session_state.df_state['Agency'].dropna().unique()
                selected_agencies = st.sidebar.multiselect("Agencija", options=sorted(all_agencies), default=sorted(all_agencies))
                all_package_types = st.session_state.df_state['Package Type'].dropna().unique()
                selected_package_types = st.sidebar.multiselect("Tip paketa", options=sorted(all_package_types), default=sorted(all_package_types))

                st.sidebar.markdown("---")
                st.sidebar.header("Filter po Datumu Putovanja")

                start_date, end_date = None, None
                if not st.session_state.df_state['Begin Date'].dropna().empty:
                    min_date = st.session_state.df_state['Begin Date'].dropna().min().date()
                    max_date = st.session_state.df_state['Begin Date'].dropna().max().date()
                    start_date = st.sidebar.date_input("Početni datum", min_date, min_value=min_date, max_value=max_date)
                    end_date = st.sidebar.date_input("Krajnji datum", max_date, min_value=min_date, max_value=max_date)
                else:
                    st.sidebar.warning("Nedostaju datumi putovanja za filtriranje.")


                col1_1, col1_2 = st.columns([0.8, 0.2])
                with col1_1:
                    st.header("📈 Analitički Dashboard")
                    if start_date and end_date:
                        st.subheader(f"Prikaz za period: {start_date.strftime('%d.%m.%Y.')} - {end_date.strftime('%d.%m.%Y.')}")

                with col1_2:
                    st.markdown("---")
                    if st.button("🔄 Ažuriraj izvještaj"):
                        st.rerun()

                df_filtered = st.session_state.df_state[st.session_state.df_state['Uključi u analizu']].copy()
                
                if selected_cities:
                    df_filtered = df_filtered[df_filtered['Arrival City'].isin(selected_cities)]
                if selected_hotels:
                    df_filtered = df_filtered[df_filtered['Hotel Name'].isin(selected_hotels)]
                if selected_authors:
                    df_filtered = df_filtered[df_filtered['Author'].isin(selected_authors)]
                if selected_agencies:
                    df_filtered = df_filtered[df_filtered['Agency'].isin(selected_agencies)]
                if selected_package_types:
                    df_filtered = df_filtered[df_filtered['Package Type'].isin(selected_package_types)]
                if start_date and end_date:
                    start_date_ts = pd.to_datetime(start_date)
                    end_date_ts = pd.to_datetime(end_date)
                    df_filtered = df_filtered[(df_filtered['Begin Date'] >= start_date_ts) & (df_filtered['Begin Date'] <= end_date_ts)]

                if df_filtered.empty:
                    st.warning("Nema podataka koji odgovaraju odabranim filterima. Molimo označite rezervacije za analizu.")
                else:
                    total_sales = df_filtered['Sale Price'].sum()
                    total_profit = df_filtered['Profit'].sum()
                    num_reservations = len(df_filtered)
                    avg_profit_per_res = total_profit / num_reservations if num_reservations > 0 else 0
                    kpi_summary = {
                        "Ukupan prihod": f"{total_sales:,.2f} BAM",
                        "Ukupan profit": f"{total_profit:,.2f} BAM",
                        "Broj rezervacija": f"{num_reservations}",
                        "Prosječan profit po rezervaciji": f"{avg_profit_per_res:,.2f} BAM"
                    }
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Ukupan Prihod (Sale Price)", f"{total_sales:,.2f} BAM")
                    col2.metric("Ukupan Profit", f"{total_profit:,.2f} BAM")
                    col3.metric("Broj Rezervacija", f"{num_reservations}")
                    col4.metric("Prosječan Profit / Rezervaciji", f"{avg_profit_per_res:,.2f} BAM")
                    st.markdown("---")

                    st.subheader("Trend Profitabilnosti Tokom Vremena")
                    profit_over_time = df_filtered.set_index('Begin Date').groupby(pd.Grouper(freq='M'))['Profit'].sum().reset_index()
                    profit_over_time['Month'] = profit_over_time['Begin Date'].dt.strftime('%Y-%m')
                    fig_profit_trend = px.line(profit_over_time, x='Month', y='Profit', title="Mjesečni profit", labels={'Month': 'Mjesec', 'Profit': 'Ukupan Profit (BAM)'}, markers=True)
                    fig_profit_trend.update_layout(xaxis_title="Mjesec", yaxis_title="Ukupan Profit (BAM)")
                    st.plotly_chart(fig_profit_trend, use_container_width=True)
                    st.markdown("---")


                    col_viz1, col_viz2 = st.columns(2)
                    with col_viz1:
                        st.subheader("Profitabilnost po Destinaciji")
                        profit_by_city = df_filtered.groupby('Arrival City')['Profit'].sum().sort_values(ascending=False).reset_index()
                        fig_city_profit = px.bar(profit_by_city.head(10), x='Arrival City', y='Profit', title="TOP 10 destinacija po profitu", text_auto='.2s', labels={'Arrival City': 'Destinacija', 'Profit': 'Ukupan Profit (BAM)'})
                        fig_city_profit.update_traces(textposition='outside')
                        st.plotly_chart(fig_city_profit, use_container_width=True)
                    with col_viz2:
                        st.subheader("Najprodavaniji Hoteli (Broj Noćenja)")
                        nights_by_hotel = df_filtered.groupby('Hotel Name')['Night'].sum().sort_values(ascending=False).reset_index()
                        fig_nights_hotel = px.bar(nights_by_hotel.head(10), x='Hotel Name', y='Night', title="TOP 10 hotela po broju noćenja", labels={'Hotel Name': 'Naziv Hotela', 'Night': 'Broj Noćenja'})
                        st.plotly_chart(fig_nights_hotel, use_container_width=True)
                    
                    col_viz3, col_viz4 = st.columns(2)
                    with col_viz3:
                        st.subheader("Udio profita po tipu paketa")
                        package_type_profit = df_filtered.groupby('Package Type')['Profit'].sum().reset_index()
                        fig_package_profit = px.pie(package_type_profit, names='Package Type', values='Profit', title="Udio profita po tipu paketa", hole=0.3)
                        st.plotly_chart(fig_package_profit, use_container_width=True)
                    with col_viz4:
                        st.subheader("Broj rezervacija po broju putnika")
                        pax_counts = df_filtered['Total Pax'].value_counts().sort_index().reset_index()
                        pax_counts.columns = ['Total Pax', 'Broj Rezervacija']
                        fig_pax_dist = px.bar(pax_counts, x='Total Pax', y='Broj Rezervacija', title="Raspodjela rezervacija po broju putnika", labels={'Total Pax': 'Broj Putnika', 'Broj Rezervacija': 'Broj Rezervacija'})
                        st.plotly_chart(fig_pax_dist, use_container_width=True)

                    col_viz5, col_viz6 = st.columns(2)
                    with col_viz5:
                        st.subheader("Analiza po Agenciji")
                        agency_profit = df_filtered.groupby('Agency')['Profit'].sum().sort_values(ascending=False).reset_index()
                        fig_agency_profit = px.bar(agency_profit.head(10), x='Agency', y='Profit', title="TOP 10 agencija po profitu", text_auto='.2s', labels={'Agency': 'Agencija', 'Profit': 'Ukupan Profit (BAM)'})
                        st.plotly_chart(fig_agency_profit, use_container_width=True)
                    with col_viz6:
                        st.subheader("Učinak po Autoru")
                        author_profit = df_filtered.groupby('Author')['Profit'].sum().sort_values(ascending=False).reset_index()
                        fig_author_profit = px.bar(author_profit.head(10), x='Author', y='Profit', title="TOP 10 autora po ostvarenom profitu", text_auto='.2s', labels={'Author': 'Autor', 'Profit': 'Ukupan Profit (BAM)'})
                        st.plotly_chart(fig_author_profit, use_container_width=True)
                    
                    st.markdown("---")
                    st.subheader("Detaljan Prikaz Svih Podataka")
                    st.warning("Ovdje možete odznačiti rezervacije za analizu ili urediti iznose. Profit se računa automatski.")
                    
                    edited_main_df = st.data_editor(
                        st.session_state.df_state,
                        key="main_data_editor",
                        column_order=('Uključi u analizu', 'Reservation No', 'Arrival City', 'Hotel Name', 'Author', 'Create Date', 'Net Price', 'Agency Amount to Pay', 'Sale Price', 'Profit'),
                        column_config={
                            "Uključi u analizu": st.column_config.CheckboxColumn("Uključi"),
                            "Profit": st.column_config.NumberColumn("Profit (BAM)", format="%.2f"),
                            "Net Price": st.column_config.NumberColumn("Neto Cijena (BAM)", format="%.2f"),
                            "Agency Amount to Pay": st.column_config.NumberColumn("Agenciji za uplatu (BAM)", format="%.2f"),
                            "Sale Price": st.column_config.NumberColumn("Prodajna Cijena (BAM)", format="%.2f"),
                            "Reservation No": st.column_config.TextColumn("Broj rezervacije"),
                            "Arrival City": "Destinacija",
                            "Hotel Name": "Hotel",
                            "Author": "Autor",
                            "Create Date": "Datum kreiranja",
                        },
                        disabled=('Reservation No', 'Arrival City', 'Hotel Name', 'Author', 'Create Date', 'Profit'),
                        use_container_width=True
                    )
                    
                    edited_main_df['Profit'] = edited_main_df['Agency Amount to Pay'] - edited_main_df['Net Price']
                    st.session_state.df_state = edited_main_df
                    
                    st.markdown("---")
                    st.subheader("Preuzimanje Analiziranog Izvještaja")
                    excel_data = to_excel(df_filtered, kpi_summary)
                    st.download_button(label="📥 Preuzmi Excel fajl", data=excel_data, file_name="analiza_rezervacija.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                st.markdown("---")
                st.subheader("🛠️ Obrada rezervacija bez cijena")
                st.warning("Pronađene su rezervacije kojima nedostaju podaci o cijeni. Unesite vrijednosti i odaberite ih za uključivanje u analizu.")
                
                df_without_prices = st.session_state.df_state[(st.session_state.df_state['Net Price'] == 0) | (st.session_state.df_state['Agency Amount to Pay'] == 0)].copy()
                
                if not df_without_prices.empty:
                    edited_missing_df = st.data_editor(
                        df_without_prices,
                        key="missing_prices_editor",
                        column_order=('Uključi u analizu', 'Reservation No', 'Arrival City', 'Hotel Name', 'Author', 'Create Date', 'Net Price', 'Agency Amount to Pay', 'Sale Price', 'Profit'),
                        column_config={
                            "Uključi u analizu": st.column_config.CheckboxColumn("Uključi", default=False),
                            "Net Price": st.column_config.NumberColumn("Neto Cijena (BAM)", format="%.2f", min_value=0),
                            "Agency Amount to Pay": st.column_config.NumberColumn("Agenciji za uplatu (BAM)", format="%.2f", min_value=0),
                            "Sale Price": st.column_config.NumberColumn("Prodajna Cijena (BAM)", format="%.2f", min_value=0),
                            "Profit": st.column_config.NumberColumn("Profit (BAM)", format="%.2f"),
                            "Reservation No": st.column_config.TextColumn("Broj rezervacije"),
                            "Arrival City": "Destinacija",
                            "Hotel Name": "Hotel",
                            "Author": "Autor",
                            "Create Date": "Datum kreiranja",
                        },
                        disabled=('Reservation No', 'Arrival City', 'Hotel Name', 'Author', 'Create Date', 'Profit'),
                        use_container_width=True
                    )
                    
                    edited_missing_df['Profit'] = edited_missing_df['Agency Amount to Pay'] - edited_missing_df['Net Price']
                    st.session_state.df_state.update(edited_missing_df)

        else:
            st.error("Nije moguće pronaći zaglavlje tabele.")
            st.warning("Molimo provjerite da li vaš Excel fajl sadrži ispravne nazive kolona u prvih 10 redova.")

    except Exception as e:
        st.error(f"Došlo je do neočekivane greške pri obradi fajla: {e}")
        st.error("Molimo provjerite da li je fajl ispravan (npr. da nije oštećen ili prazan).")
