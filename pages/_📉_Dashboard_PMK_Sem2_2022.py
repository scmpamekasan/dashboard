from ast import Return
from functools import cache
import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit
import matplotlib.pyplot as plt
import plotly.figure_factory as ff
import streamlit.components.v1 as components
import numpy as np
import altair as alt
from PIL import Image
import pydeck as pdk
import plotly.graph_objects as go
import xlsxwriter
from matplotlib.patches import ConnectionPatch
from streamlit_option_menu import option_menu

st.set_page_config(layout="wide")

################################### HEADER ###################################
st.markdown(""" <style> .font {font-size:50px ; text-align: center; font-family: 'Cooper Black'; color: #A70D2A;} </style> """, unsafe_allow_html=True)
st.markdown('<p class="font">Semester 2 - 2022</p>', unsafe_allow_html=True)

########## Menu SideBar #############

choose = option_menu("", ["DSO Sales Performance", "DSO Canvas Performance", "KPI Realisasi Kanvas-TDN-EC", "Rank Promotor"],
                         icons=['graph-up-arrow', 'bar-chart-line-fill', 'graph-up', 'award-fill'],
                          default_index=0,
                         styles={
                            "container": {"padding": "5!important", "background-color": "#fafafa"},
                            "icon": {"color": "orange", "font-size": "22px"}, 
                            "nav-link": {"font-size": "14px", "text-align": "left", "margin":"0px", "--hover-color": "#eee"},
                            "nav-link-selected": {"background-color": "#990000"},
                            },
                         orientation = 'horizontal'
                            )

########## Menu SideBar #############

###option = st.sidebar.selectbox("Menu",[
###                              "DSO Sales Performance",
###                              "DSO Canvas Performance",
###                              "KPI Realisasi Kanvas-TDN-EC",
###                              "Rank Promotor"
###                              ]
###                              )

################ Home Page #############
###if option == "Home Page":
###    st.markdown("""---""")
###
###    height = 200
###    width = 1200
###    
###    df = pd.DataFrame(
###        np.random.randn(1, 2) / [50, 50] + [(-7.004320, 113.861274), (-7.187682, 113.240499),(-7.160772, 113.482642), (-7.048228, 112.781668)],
###        columns=['lat', 'lon'],)
###
###    st.pydeck_chart(pdk.Deck(
###        map_style='mapbox://styles/mapbox/light-v9',
###        initial_view_state=pdk.ViewState(
###            latitude=-7.060772,
###            longitude=113.340499,
###            zoom=8,
###            height=height,
###            width=width
###        ),
###        layers=[
###            pdk.Layer(
###                'ScatterplotLayer',
###                data=df,
###                get_position='[lon, lat]',
###                get_color='[200, 30, 0, 160]',
###                get_radius=200,
###                auto_highlight=True
###            ),
###        ],
###    ))
###
###    ####################### MAIN PAGE ####################
###
###    col1, col2, col3, col4 = st.columns(4)
###    col1.metric("MP Jt Batang/Hari", "12,4")
###    col2.metric("OU", "48.560")
###    col3.metric("Total Penduduk", "4.004.563")
###    col4.metric("Jumlah Zona", "4")
###
###
###    col1, col2, col3, col4 = st.columns(4)
###    col1.metric("MP Ball Standard/Mgg", "30.880")
###    col2.metric("Outlet Database", "9.413")
###    col3.metric("Laki Laki (Smokers)", "1.212.375")
###    col4.metric("Jumlah Sektor", "72")
###
###    st.markdown("""---""")
###    Zona = st.radio(
###        "Select Zona",
###        ('Bangkalan', 'Sampang', 'Pamekasan', 'Sumenep'), horizontal= True)
###
###    if Zona == 'Bangkalan':
###        col1, col2, col3, col4 = st.columns(4)
###        col1.metric("MP Jt Batang/Hari", "3,3")
###        col1.metric("MP Ball Standard/Mgg", "8.136")
###        col2.metric("OU", "12.541")
###        col2.metric("Outlet Database", "2.580")
###        col3.metric("Total Penduduk", "1.034.197")
###        col3.metric("Laki Laki (Smokers)", "319.431")
###        col4.metric("Jumlah Sektor", "17")
###        col4.metric("Perkapita Rokok/Mgg 2021 (Rp)", "18.994")
###
###    if Zona == 'Sampang':
###        col1, col2, col3, col4 = st.columns(4)
###        col1.metric("MP Jt Batang/hari", "3,1")
###        col1.metric("MP Ball Standard/Mgg", "7.871")
###        col2.metric("OU", "12.076")
###        col2.metric("Outlet Database", "2.100")
###        col3.metric("Total Penduduk", "995.874")
###        col3.metric("Laki Laki (Smokers)", "309.025")
###        col4.metric("Jumlah Sektor", "15")
###        col4.metric("Perkapita Rokok/Mgg (Rp)", "17.521")
###
###    if Zona == 'Pamekasan':
###        col1, col2, col3, col4 = st.columns(4)
###        col1.metric("MP Jt Batang/Hari", "2,6")
###        col1.metric("MP Ball Standard/Mgg", "6.506")
###        col2.metric("OU", "10.308")
###        col2.metric("Outlet Database", "2.240")
###        col3.metric("Total Penduduk", "850.057")
###        col3.metric(" Laki Laki (Smokers)", "255.437")
###        col4.metric("Jumlah Sektor", "13")
###        col4.metric("Perkapita Rokok/Mgg (Rp)", "17.599")
###
###    if Zona == 'Sumenep' :
###        col1, col2, col3, col4 = st.columns(4)
###    
###        col1.metric("MP Jt Batang/Hari", "3,3")
###        col1.metric("MP Ball Standard/Mgg", "8.367")
###        col2.metric("OU", "13.635")
###        col2.metric("Outlet Database", "2.493")
###        col3.metric("Total Penduduk", "1.124.436")
###        col3.metric(" Laki Laki (Smokers)", "328.483")
###        col4.metric("Jumlah Sektor", "27")
###        col4.metric("Perkapita Rokok/Mgg (Rp)", "16.719")
###
################## Struktur Organisasi ###############
###
###if option == "Struktur Organisasi":
###    st.markdown("""---""")
###    st.header("Struktur Organisasi")
###    col1, col2, col3, col4 = st.columns(4)
###    with col1 :
###        Team = st.selectbox("Nama Team",["All Team","Promotion Team", "Distribution Team", "Back Office Team"])
###
###    if Team == "All Team" :
###        image1 = Image.open('/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/Team_Promosi.png')
###        image2 = Image.open('/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/Team_Distribusi.png')
###        image3 = Image.open('/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/Back_Office.png')
###
###        col1, col2, col3 = st.columns(3)
###        with col1 :
###                st.subheader('Promotion Team')
###                st.image(image1, caption='Promotion Team')
###        with col2 :
###                st.subheader('Distribution Team')
###                st.image(image2, caption='Distribution Team')
###        with col3 :
###                st.subheader('Back Office')
###                st.image(image3, caption='Back Office')
###
###    if Team == "Promotion Team" :
###        st.subheader('Promotion Team')
###        col1, col2, col3, col4 = st.columns(4)
###        col1.metric("Team Kanvas", "14")
###        col2.metric("Armada Kanvas", "14")
###        col3.metric("Team Khusus", "3")
###        col4.metric("Armada Team Khusus", "3")     
###
###        col1, col2 = st.columns(2)
###        df = pd.read_excel(
###            io ="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan//Map Territory Juli 2022 DSO Pamekasan.xlsx",
###            engine="openpyxl",
###            sheet_name="Team_Kanvas",
###            usecols="C:H",
###            nrows=15,
###            )
###        
###        df2 = pd.read_excel(
###            io ="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/Map Territory Juli 2022 DSO Pamekasan.xlsx",
###            engine="openpyxl",
###            sheet_name="Team_Khusus",
###            usecols="C:G",
###            nrows=15,
###            )
###        with col1 :
###            df1 = df.loc[:,["APR/PR", "Team Kanvas", "Tipe Kendaraan","No Polisi"]]
###            df1.set_index("Team Kanvas", inplace=True)
###            st. write(df1)
###
###        with col2 :
###            df3 = df2.loc[:,["APR/PR","Team Khusus","Tipe Kendaraan","No Polisi"]].head(3)
###            df3.set_index("Team Khusus", inplace=True)
###            st. write(df3)
###
###    if Team == "Distribution Team" :
###        st.subheader('Distribution Team')
###        col1, col2, col3, col4 = st.columns(4)
###        col1.metric("Team Shop", "7")
###        col2.metric("Armada Shop", "7")
###        col3.metric("Team SR/MR - MMC", "3")
###        col4.metric("Armada Team SR/MR - MMC", "3")  
###        
###
###
################ MP By Zona - Kategori - Kelas #############
###if option == "MP By Zona - Kategori - Kelas":
###    st.markdown("""---""")
###
###    st.markdown('### MP By ZONA')
###    
###    @st.cache(allow_output_mutation=True)
###    def load_data_1():
###        df = pd.read_excel(
###            io="/Users/ekasulawestara/Desktop/Dashboard_PMK_2022/st-multi_app/pages/MIT.xlsx",
###            engine="openpyxl",
###            sheet_name="MP_Zona",
###            usecols="H:L",
###            nrows=100,
###            )
###        return df
###
###    df = load_data_1()
###
###    fig = px.pie(df, values='MP_Ball_Standard', names='Zona',width= 500, height= 400, title="MP By Zona Per Jun 2022")
###    fig.update_traces(textposition='inside', textinfo='percent+label')
###                                            
###    st.plotly_chart(fig) 
###
###    st.markdown("""---""")
###
###    st.markdown('### MP DSO By Kategori & Kelas')
###    col1, col2 = st.columns(2)
###    with col2 :
###        if option == "MP By Zona - Kategori - Kelas":
######            @st.cache(allow_output_mutation=True)      <<< Masih gagal pakai Cache
######            def load_data_2():    <<< Masih gagal pakai Cache
###                df = pd.read_excel(
###                    io="/Users/ekasulawestara/Desktop/Dashboard_PMK_2022/st-multi_app/pages/MIT.xlsx",
###                    engine="openpyxl",
###                    sheet_name="Sektor",
###                    usecols="A:P",
###                    nrows=10000,
###                    )
######                return df   <<< Masih gagal pakai Cache
######            df = load_data_2   <<< Masih gagal pakai Cache
###                
###                fig = px.pie(df, values='Jun22', names='Kategori-Kelas',width= 500, height= 400, title="MP By Kategori Rokok (Kelas) Per Jun 22")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###
###                st.plotly_chart(fig)
###
###    with col1 :
###        if option == "MP By Zona - Kategori - Kelas":
######            @st.cache(allow_output_mutation=True)   <<< Masih gagal pakai Cache
######            def load_data_3():   <<< Masih gagal pakai Cache
###                df = pd.read_excel(
###                    io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/Market_DSO.xlsx",
###                    engine="openpyxl",
###                    sheet_name="MARKET",
###                    usecols="B:C",
###                    nrows=1000,
###                    )
######                return df   <<< Masih gagal pakai Cache
######            df = load_data_3   <<< Masih gagal pakai Cache
###                
###                fig = px.pie(df, values='MP', names='M.Potensi_per_Kategori_Rokok',width= 500, height= 400, title="MP By Kategori Rokok 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###                st.plotly_chart(fig)
###
###    
###    @st.cache(allow_output_mutation=True)
###    def load_data_4():
###        df = pd.read_excel(
###            io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/MIT.xlsx",
###            engine="openpyxl",
###            sheet_name="MP_Zona",
###            usecols="H:L",
###            nrows=48,
###            )
###        return df
###
###    df = load_data_4()
###
###    col1, col2 = st.columns(2)
###    with col1 :
###        if option == "MP By Zona - Kategori - Kelas":
###
###            Zona = st.radio(
###                "Select Zona",
###                ('Bangkalan', 'Sampang', 'Pamekasan', 'Sumenep'), horizontal= True)
###            if Zona == 'Bangkalan':
###                
###                table = pd.pivot_table( data=df, 
###                                        index=['Kategori'], 
###                                        columns=['Zona'], 
###                                        values='MP_Ball_Standard',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###                ###table.loc[:,Zona]
###
###                fig = px.pie(table, values=Zona, names= ['SKM', 'SKML','SKT','SPM'], width= 500, height= 400, title="MP Zona Bangkalan By Kategori Per Jun 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###
###                st.plotly_chart(fig)
###    
###                
###
###            if Zona == 'Sampang':
###                table = pd.pivot_table( data=df, 
###                                        index=['Kategori'], 
###                                        columns=['Zona'], 
###                                        values='MP_Ball_Standard',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###            
###                ###table.loc[:,Zona]
###                
###                fig = px.pie(table, values=Zona,names= ['SKM', 'SKML','SKT','SPM'],width= 500, height= 400, title="MP Zona Sampang By Kategori Per Jun 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###
###                st.plotly_chart(fig)
###
###            if Zona == 'Pamekasan':
###                table = pd.pivot_table( data=df, 
###                                        index=['Kategori'], 
###                                        columns=['Zona'], 
###                                        values='MP_Ball_Standard',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###            
###                ###table.loc[:,Zona]
###                
###                fig = px.pie(table, values=Zona, names= ['SKM', 'SKML','SKT','SPM'], width= 500, height= 400, title="MP Zona Pamekasan By Kategori Per Jun 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###
###                st.plotly_chart(fig)
###
###            if Zona == 'Sumenep' :
###                table = pd.pivot_table( data=df, 
###                                        index=['Kategori'], 
###                                        columns=['Zona'], 
###                                        values='MP_Ball_Standard',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###            
###                ###table.loc[:,Zona]
###                
###                fig = px.pie(table, values=Zona, names= ['SKM', 'SKML','SKT','SPM'], width= 500, height= 400, title="MP Zona Sumenep By Kategori Per Jun  2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###
###                st.plotly_chart(fig)
###                
###                
###    with col2 :
###        if option == "MP By Zona - Kategori - Kelas":
###            st.markdown("""---""")
###            
###            if Zona == 'Bangkalan':
###                table = pd.pivot_table( data=df, 
###                                        index=['Kat-Kel'], 
###                                        columns=['Zona'], 
###                                        values='MP_Ball_Standard',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###            
###                fig = px.pie(table, values=Zona, names= ['SKM-Low','SKM-Med','SKM-Prem','SKML-Low','SKML-Med','SKML-Prem','SKT-Low','SKT-Med','SKT-Prem','SPM-Low','SPM-Med','SPM-Prem'], width= 500, height= 400, title="MP Zona Bangkalan By Kategori-Kelas Per Jun 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###            
###                st.plotly_chart(fig)
###
###            if Zona == 'Sampang':
###                table = pd.pivot_table( data=df, 
###                                        index=['Kat-Kel'], 
###                                        columns=['Zona'], 
###                                        values='MP_Ball_Standard',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###            
###                fig = px.pie(table, values=Zona, names= ['SKM-Low','SKM-Med','SKM-Prem','SKML-Low','SKML-Med','SKML-Prem','SKT-Low','SKT-Med','SKT-Prem','SPM-Low','SPM-Med','SPM-Prem'], width= 500, height= 400, title="MP Zona Sampang By Kategori-Kelas Per Jun 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###    
###                st.plotly_chart(fig)
###
###
###            if Zona == 'Pamekasan':
###                table = pd.pivot_table( data=df, 
###                                        index=['Kat-Kel'], 
###                                        columns=['Zona'], 
###                                        values='MP_Ball_Standard',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###            
###                fig = px.pie(table, values=Zona, names= ['SKM-Low','SKM-Med','SKM-Prem','SKML-Low','SKML-Med','SKML-Prem','SKT-Low','SKT-Med','SKT-Prem','SPM-Low','SPM-Med','SPM-Prem'], width= 500, height= 400, title="MP Zona Pamekasan By Kategori-Kelas Per Jun 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###    
###                st.plotly_chart(fig)
###
###
###            if Zona == 'Sumenep' :
###                table = pd.pivot_table( data=df, 
###                                        index=['Kat-Kel'], 
###                                        columns=['Zona'], 
###                                        values='MP_Ball_Standard',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###            
###                fig = px.pie(table, values=Zona, names= ['SKM-Low','SKM-Med','SKM-Prem','SKML-Low','SKML-Med','SKML-Prem','SKT-Low','SKT-Med','SKT-Prem','SPM-Low','SPM-Med','SPM-Prem'], width= 500, height= 400, title="MP Zona Sumenep By Kategori-Kelas Per Jun 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###    
###                st.plotly_chart(fig)
###
################ MS By Group Pabrikan #############
###
###if option == "MS By Group Pabrikan":
###    col1, col2 = st.columns(2)
###    with col1 :
###        st.markdown('### MS By Group - Nielsen (Per Mei 2022)')
###        fig = plt.figure(figsize=(9, 5.0625))
###        ax1 = fig.add_subplot(121)
###        ax2 = fig.add_subplot(122)
###        fig.subplots_adjust(wspace=0)
###        # large pie chart parameters
###
###        df = pd.read_excel(
###                io="/Users/ekasulawestara/Desktop/Dashboard_PMK_2022/st-multi_app/pages/Market_DSO.xlsx",
###                engine="openpyxl",
###                sheet_name="Nielsen",
###                usecols="A:P",
###                nrows=1000,
###                )
###
###        pivot1 = pd.pivot_table(data=df, index=['Group'], values = 'May-22', aggfunc='sum')
###        xxx = [pivot1.iloc[0]['May-22'], pivot1.iloc[1]['May-22'], pivot1.iloc[2]['May-22'], pivot1.iloc[3]['May-22'], pivot1.iloc[4]['May-22'],pivot1.iloc[5]['May-22']]
###        pivot2 = pd.pivot_table(data=df, index=['Group', 'Kategori'], values = 'May-22', aggfunc='sum')
###        yyy = [pivot2.iloc[4]['May-22'], pivot2.iloc[5]['May-22'], pivot2.iloc[6]['May-22']]
###
###        ratios = xxx
###        labels = ['BAT', 'Djarum', 'GG', 'Nojorono', 'Others', 'PMI']
###        explode = [0, 0.1, 0, 0, 0, 0]
###        # rotate so that first wedge is split by the x-axis
###        angle = -60 * ratios[0]
###        ax1.pie(ratios, autopct='%1.1f%%', startangle=angle,
###                labels=labels, explode=explode)
###        # small pie chart parameters
###        ratios = yyy
###        labels = ['SKM', 'SKML', 'SKT']
###        width = .2
###        ax2.pie(ratios, autopct='%1.1f%%', startangle=angle,
###                labels=labels, radius=0.5, textprops={'size': 'smaller'})
###        ax1.set_title('MS by Group')
###        ax2.set_title('Kontribusi by Kategori')
###        # use ConnectionPatch to draw lines between the two plots
###        # get the wedge data
###        theta1, theta2 = ax1.patches[1].theta1, ax1.patches[1].theta2
###        center, r = ax1.patches[1].center, ax1.patches[1].r
###        # draw top connecting line
###        x = r * np.cos(np.pi / 180 * theta2) + center[0]
###        y = np.sin(np.pi / 180 * theta2) + center[1]
###        con = ConnectionPatch(xyA=(- width / 2, .5), xyB=(x, y),
###                            coordsA="data", coordsB="data", axesA=ax2, axesB=ax1)
###        con.set_color([0, 0, 0])
###        con.set_linewidth(0.5)
###        ax2.add_artist(con)
###        # draw bottom connecting line
###        x = r * np.cos(np.pi / 180 * theta1) + center[0]
###        y = np.sin(np.pi / 180 * theta1) + center[1]
###        con = ConnectionPatch(xyA=(- width / 2, -.5), xyB=(x, y), coordsA="data",
###                            coordsB="data", axesA=ax2, axesB=ax1)
###        con.set_color([0, 0, 0])
###        ax2.add_artist(con)
###        con.set_linewidth(0.5)
###        st.write(fig)
###
###
###    with col2 :    
###
###        st.markdown('### MS By Group - Internal (Per Juni 2022)')
###        fig = plt.figure(figsize=(9, 5.0625))
###        ax1 = fig.add_subplot(121)
###        ax2 = fig.add_subplot(122)
###        fig.subplots_adjust(wspace=0)
###        
###        # large pie chart parameters
###        
###        @st.cache(allow_output_mutation=True)
###        def load_data_6():
###            df = pd.read_excel(
###            io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/MIT.xlsx",
###            engine="openpyxl",
###            sheet_name="DSO",
###            usecols="A:R",
###            nrows=200,
###            )
###            return df
###
###        df = load_data_6()
###        
###        pivot1 = pd.pivot_table(data=df, index=['Group'], values = 'Jun22', aggfunc='sum')
###        xxx1 = [pivot1.iloc[0]['Jun22'], pivot1.iloc[1]['Jun22'], pivot1.iloc[2]['Jun22'], pivot1.iloc[3]['Jun22'], pivot1.iloc[4]['Jun22'], pivot1.iloc[5]['Jun22']]
###        pivot2 = pd.pivot_table(data=df, index=['Kategori', 'Group'], values = 'Jun22', aggfunc='sum')
###        yyy1 = [pivot2.iloc[1]['Jun22'], pivot2.iloc[6]['Jun22'], pivot2.iloc[11]['Jun22']]
###        ratios = xxx1
###        labels = ['NTI', 'Djarum', 'GG', 'BAT', 'Others', 'PMI']
###        explode = [0, 0.1, 0, 0, 0, 0]
###        # rotate so that first wedge is split by the x-axis
###        angle = -33 * ratios[0]
###        ax1.pie(ratios, autopct='%1.1f%%', startangle=angle,
###                labels=labels, explode=explode)
###        # small pie chart parameters
###        ratios = yyy1
###        labels = ['SKM', 'SKML', 'SKT']
###        width = .2
###        ax2.pie(ratios, autopct='%1.1f%%', startangle=angle,
###                labels=labels, radius=0.5, textprops={'size': 'smaller'})
###        ax1.set_title('MS by Group')
###        ax2.set_title('Kontribusi by Kategori')
###        # use ConnectionPatch to draw lines between the two plots
###        # get the wedge data
###        theta1, theta2 = ax1.patches[1].theta1, ax1.patches[1].theta2
###        center, r = ax1.patches[1].center, ax1.patches[1].r
###        # draw top connecting line
###        x = r * np.cos(np.pi / 180 * theta2) + center[0]
###        y = np.sin(np.pi / 180 * theta2) + center[1]
###        con = ConnectionPatch(xyA=(- width / 2, .5), xyB=(x, y),
###                            coordsA="data", coordsB="data", axesA=ax2, axesB=ax1)
###        con.set_color([0, 0, 0])
###        con.set_linewidth(.5)
###        ax2.add_artist(con)
###        # draw bottom connecting line
###        x = r * np.cos(np.pi / 180 * theta1) + center[0]
###        y = np.sin(np.pi / 180 * theta1) + center[1]
###        con = ConnectionPatch(xyA=(- width / 2, -.5), xyB=(x, y), coordsA="data",
###                            coordsB="data", axesA=ax2, axesB=ax1)
###        con.set_color([0, 0, 0])
###        ax2.add_artist(con)
###        con.set_linewidth(.5)
###        st.write(fig)
###
################# MS Zona By Group ##############
###if option == "MS By Group Pabrikan":
###        Zona = st.radio(
###            "Select Zona",
###            ('All Zona', 'Bangkalan', 'Sampang', 'Pamekasan', 'Sumenep'), horizontal= True)
###
###        df_dso = pd.read_excel(
###            io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/MIT.xlsx",
###            engine="openpyxl",
###            sheet_name="DSO",
###            usecols="A:T",
###            nrows=200,
###            )
###    
###        if Zona == 'All Zona':
###                
###                table_dso = pd.pivot_table( data=df_dso, 
###                                        index=['Group','Kategori'], 
###                                        columns=['DSO'], 
###                                        values='Jun22',
###                                        aggfunc='sum',
###                                        fill_value=0
###                                                )
###                ###table_dso.loc[:,Zona]
###                fig = px.pie(table_dso, values='Pamekasan', names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by All Zona Per Jun 2022")
###                fig.update_traces(textposition='inside', textinfo='percent+label')
###                st.plotly_chart(fig)
###
###
###        df = pd.read_excel(
###            io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/MIT.xlsx",
###            engine="openpyxl",
###            sheet_name="Zona",
###            usecols="A:N",
###            nrows=500,
###            )
###
###        if Zona == 'Bangkalan':
###
###                    table = pd.pivot_table( data=df, 
###                                            index=['Group','Kategori'], 
###                                            columns=['Zona'], 
###                                            values='Jun22',
###                                            aggfunc='sum',
###                                            fill_value=0
###                                                    )
###                    ###table.loc[:,Zona]
###
###                    fig = px.pie(table, values=Zona, names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by Zona Bangkalan Per Jun 2022")
###                    fig.update_traces(textposition='inside', textinfo='percent+label')
###                    st.plotly_chart(fig)
###
###        if Zona == 'Sampang':
###                    table = pd.pivot_table( data=df, 
###                                            index=['Group','Kategori'], 
###                                            columns=['Zona'], 
###                                            values='Jun22',
###                                            aggfunc='sum',
###                                            fill_value=0
###                                                    )
###                    ###table.loc[:,Zona]
###
###                    fig = px.pie(table, values=Zona, names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by Zona Sampang Per Jun 2022")
###                    fig.update_traces(textposition='inside', textinfo='percent+label')
###                    st.plotly_chart(fig)
###
###        if Zona == 'Pamekasan':
###                    table = pd.pivot_table( data=df, 
###                                            index=['Group','Kategori'], 
###                                            columns=['Zona'], 
###                                            values='Jun22',
###                                            aggfunc='sum',
###                                            fill_value=0
###                                                    )
###                    ###table.loc[:,Zona]
###
###                    fig = px.pie(table, values=Zona, names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by Zona Pamekasan Per Jun 2022")
###                    fig.update_traces(textposition='inside', textinfo='percent+label')
###                    st.plotly_chart(fig)
###
###        if Zona == 'Sumenep' :
###                    table = pd.pivot_table( data=df, 
###                                            index=['Group','Kategori'], 
###                                            columns=['Zona'], 
###                                            values='Jun22',
###                                            aggfunc='sum',
###                                            fill_value=0
###                                                    )
###                    ###table.loc[:,Zona]
###
###                    fig = px.pie(table, values=Zona, names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by Zona Sumenep Per Jun 2022")
###                    fig.update_traces(textposition='inside', textinfo='percent+label')
###                    st.plotly_chart(fig)
###        
###        ######### MS Zona By Group Kategori ########
###
###        if st.button ("Show Detail"):  
###
###            if Zona =='All Zona':
###                col1, col2 = st.columns(2)
###                with col1 :        
###                        table_dso = pd.pivot_table( data=df_dso, 
###                                                index=['Group-Kat'], 
###                                                columns=['DSO'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table_dso, values='Pamekasan', names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by All Zona Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###                with col2 :
###                        table_dso = pd.pivot_table( data=df_dso, 
###                                                index=['Group-Kat-Kel'], 
###                                                columns=['DSO'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table_dso, values='Pamekasan', names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by All Zona Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###
###        
###            if Zona == 'Bangkalan':
###
###                col1, col2 = st.columns(2)
###                with col1 :        
###                        table = pd.pivot_table( data=df, 
###                                                index=['Group-Kat'], 
###                                                columns=['Zona'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by Zona Bangkalan Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###                with col2 :
###                        table = pd.pivot_table( data=df, 
###                                                index=['Group-Kat-Kel'], 
###                                                columns=['Zona'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by Zona Bangkalan Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###
###            if Zona == 'Sampang':
###
###                col1, col2 = st.columns(2)
###                with col1 :  
###                        table = pd.pivot_table( data=df, 
###                                                index=['Group-Kat'], 
###                                                columns=['Zona'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by Zona Sampang Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###                with col2 :
###                        table = pd.pivot_table( data=df, 
###                                                index=['Group-Kat-Kel'], 
###                                                columns=['Zona'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by Zona Sampang Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###            if Zona == 'Pamekasan':
###
###                col1, col2 = st.columns(2)
###                with col1 :  
###                        table = pd.pivot_table( data=df, 
###                                                index=['Group-Kat'], 
###                                                columns=['Zona'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by Zona Pamekasan Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###                with col2 :
###                        table = pd.pivot_table( data=df, 
###                                                index=['Group-Kat-Kel'], 
###                                                columns=['Zona'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by Zona Pamekasan Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###            if Zona == 'Sumenep' :
###
###                col1, col2 = st.columns(2)
###                with col1 :  
###                        table = pd.pivot_table( data=df, 
###                                                index=['Group-Kat'], 
###                                                columns=['Zona'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by Zona Sumenep Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###                with col2 :
###                        table = pd.pivot_table( data=df, 
###                                                index=['Group-Kat-Kel'], 
###                                                columns=['Zona'], 
###                                                values='Jun22',
###                                                aggfunc='sum',
###                                                fill_value=0
###                                                        )
###                        ###table.loc[:,Zona]
###
###                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by Zona Sumenep Per Jun 2022")
###                        fig.update_traces(textposition='inside', textinfo='percent+label')
###                        st.plotly_chart(fig)
###
###            st.button ("Close Detail")
###
############### Top 10 Brand Nielsen #############
###
###if option == "Top 10 Brand Nielsen" :
###
###    st.markdown('### Top 10 Brand MS Nielsen - Mei 2022 (YTD)')
###
###    df = pd.read_excel(
###    io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/Market_DSO.xlsx",
###    engine="openpyxl",
###    sheet_name="Nielsen",
###    usecols="A:Q",
###    nrows=10
###    )
###    new_df = df.drop(['Group','Kategori'], axis=1)
###
###    first =((new_df).iloc[0,0])
###    first_value = ('{:.2f}%'.format((new_df).iloc[0,13]))
###    first_delta = ('{:.2f}%'.format((new_df).iloc[0,14])) 
###
###    second =((new_df).iloc[1,0])
###    second_value = ('{:.2f}%'.format((new_df).iloc[1,13]))
###    second_delta = ('{:.2f}%'.format((new_df).iloc[1,14]))  
###
###    thirth =((new_df).iloc[2,0])
###    thirth_value = ('{:.2f}%'.format((new_df).iloc[2,13]))
###    thirth_delta = ('{:.2f}%'.format((new_df).iloc[2,14]))
###
###    fourth =((new_df).iloc[3,0])
###    fourth_value = ('{:.2f}%'.format((new_df).iloc[3,13]))
###    fourth_delta = ('{:.2f}%'.format((new_df).iloc[3,14]))
###
###    fifth =((new_df).iloc[4,0])
###    fifth_value = ('{:.2f}%'.format((new_df).iloc[4,13]))
###    fifth_delta = ('{:.2f}%'.format((new_df).iloc[4,14]))
###
###    col1, col2, col3, col4, col5 = st.columns(5)
###    col1.metric(first,first_value, first_delta)
###    col2.metric(second,second_value, second_delta)
###    col3.metric(thirth, thirth_value, thirth_delta)
###    col4.metric(fourth, fourth_value, fourth_delta)
###    col5.metric(fifth, fifth_value, fifth_delta)
###
###    sixth =((new_df).iloc[5,0])
###    sixth_value = ('{:.2f}%'.format((new_df).iloc[5,13]))
###    sixth_delta = ('{:.2f}%'.format((new_df).iloc[5,14])) 
###    seventh =((new_df).iloc[6,0])
###    seventh_value = ('{:.2f}%'.format((new_df).iloc[6,13]))
###    seventh_delta = ('{:.2f}%'.format((new_df).iloc[6,14]))  
###    eighth =((new_df).iloc[7,0])
###    eighth_value = ('{:.2f}%'.format((new_df).iloc[7,13]))
###    eighth_delta = ('{:.2f}%'.format((new_df).iloc[7,14]))
###    ninth =((new_df).iloc[8,0])
###    ninth_value = ('{:.2f}%'.format((new_df).iloc[8,13]))
###    ninth_delta = ('{:.2f}%'.format((new_df).iloc[8,14]))
###    tenth =((new_df).iloc[9,0])
###    tenth_value = ('{:.2f}%'.format((new_df).iloc[9,13]))
###    tenth_delta = ('{:.2f}%'.format((new_df).iloc[9,14]))
###
###    col6, col7, col8, col9, col10 = st.columns(5)
###    col6.metric(sixth,sixth_value, sixth_delta)
###    col7.metric(seventh,seventh_value, seventh_delta)
###    col8.metric(eighth, eighth_value, eighth_delta)
###    col9.metric(ninth, ninth_value, ninth_delta)
###    col10.metric(tenth, tenth_value, tenth_delta)
###
###    st.markdown("""---""")
###
###    with st.expander("Show Data (Magnificent 7 Brand Prioritas - Djarum)"):
###
###        names = ['LA Bold', 'Geo Mild', 'Chief Filter','L.A. Lights Regular','D. 7 6', 'D. Super','Chief Kretek']
###
###        df2 = pd.read_excel(
###        io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/Market_DSO.xlsx",
###        engine="openpyxl",
###        sheet_name="Nielsen",
###        usecols="A:Q",
###        nrows=100
###        )
###        df_djarum_group = df2[df2.Brand.isin(names)]
###        df_top_10_internal = df_djarum_group.drop(['Group','Kategori','Delta',], axis=1)
###        st.write(df_top_10_internal)
###
############### Top 10 Brand Internal ############
###
###if option == "Top 10 Brand Internal" :
###
###        st.markdown('### Top 10 Brand MS Internal - Juni 2022 (YTD)')
###        
###        names = ['LA Bold', 'Geo Mild', 'Chief Filter','L.A. Lights Regular','D. 7 6', 'D. Super','Chief Kretek']
###
###        df = pd.read_excel(
###        io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/MIT.xlsx",
###        engine="openpyxl",
###        sheet_name="DSO",
###        usecols="A:R",
###        nrows=10
###        )
###        new_df = df.drop(['Group','Kategori', 'Kelas','DSO','Group-Kat','Group-Kat-Kel','Agt22','Okt22', 'Des22'], axis=1)
###
###        first =((new_df).iloc[0,0])
###        first_value = ('{:.2f}%'.format((new_df).iloc[0,7]))
###        first_delta = ('{:.2f}%'.format((new_df).iloc[0,8])) 
###
###        second =((new_df).iloc[1,0])
###        second_value = ('{:.2f}%'.format((new_df).iloc[1,7]))
###        second_delta = ('{:.2f}%'.format((new_df).iloc[1,8]))  
###
###        thirth =((new_df).iloc[2,0])
###        thirth_value = ('{:.2f}%'.format((new_df).iloc[2,7]))
###        thirth_delta = ('{:.2f}%'.format((new_df).iloc[2,8]))
###
###        fourth =((new_df).iloc[3,0])
###        fourth_value = ('{:.2f}%'.format((new_df).iloc[3,7]))
###        fourth_delta = ('{:.2f}%'.format((new_df).iloc[3,8]))
###
###        fifth =((new_df).iloc[4,0])
###        fifth_value = ('{:.2f}%'.format((new_df).iloc[4,7]))
###        fifth_delta = ('{:.2f}%'.format((new_df).iloc[4,8]))
###
###        col1, col2, col3, col4, col5 = st.columns(5)
###        col1.metric(first,first_value, first_delta)
###        col2.metric(second,second_value, second_delta)
###        col3.metric(thirth, thirth_value, thirth_delta)
###        col4.metric(fourth, fourth_value, fourth_delta)
###        col5.metric(fifth, fifth_value, fifth_delta)
###
###        sixth =((new_df).iloc[5,0])
###        sixth_value = ('{:.2f}%'.format((new_df).iloc[5,7]))
###        sixth_delta = ('{:.2f}%'.format((new_df).iloc[5,8])) 
###        seventh =((new_df).iloc[6,0])
###        seventh_value = ('{:.2f}%'.format((new_df).iloc[6,7]))
###        seventh_delta = ('{:.2f}%'.format((new_df).iloc[6,8]))  
###        eighth =((new_df).iloc[7,0])
###        eighth_value = ('{:.2f}%'.format((new_df).iloc[7,7]))
###        eighth_delta = ('{:.2f}%'.format((new_df).iloc[7,8]))
###        ninth =((new_df).iloc[8,0])
###        ninth_value = ('{:.2f}%'.format((new_df).iloc[8,7]))
###        ninth_delta = ('{:.2f}%'.format((new_df).iloc[8,8]))
###        tenth =((new_df).iloc[9,0])
###        tenth_value = ('{:.2f}%'.format((new_df).iloc[9,7]))
###        tenth_delta = ('{:.2f}%'.format((new_df).iloc[9,8]))
###
###        col6, col7, col8, col9, col10 = st.columns(5)
###        col6.metric(sixth,sixth_value, sixth_delta)
###        col7.metric(seventh,seventh_value, seventh_delta)
###        col8.metric(eighth, eighth_value, eighth_delta)
###        col9.metric(ninth, ninth_value, ninth_delta)
###        col10.metric(tenth, tenth_value, tenth_delta)
###
###        st.markdown("""---""")
###        with st.expander("Show Data (Magnificent 7 Brand Prioritas - Djarum)"):
###
###            df2 = pd.read_excel(
###            io="/Users/ekasulawestara/Dashboard Sem2 - DSO Pamekasan/MIT.xlsx",
###            engine="openpyxl",
###            sheet_name="DSO",
###            usecols="A:T",
###            nrows=45
###            )
###
###            df_djarum_group = df2[df2.GabBrand.isin(names)]
###            df_top_10_internal = df_djarum_group.drop(['DSO','Group','Kategori','Group-Kat', 'Group-Kat-Kel', 'Kelas','Delta', 'Agt22', 'Okt22', 'Des22'], axis=1)
###            st.write(df_top_10_internal)
st.markdown("""___""")

############ DSO Sales Performance #############

if choose == "DSO Sales Performance":
    st.title(" 📈 Sales Performance")
    ###@st.cache(allow_output_mutation=True)
    def load_data_7():
        dfOmset_Sem2_2022 = pd.read_excel(
            io="Omset_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Omset_Sem2_2022",
            usecols="G:I",
            nrows=250000
            )
        return (dfOmset_Sem2_2022)
    dfOmset_Sem2_2022=load_data_7()
###    st.write(dfOmset_Sem2_2022)

    df11 = dfOmset_Sem2_2022.pivot_table('Omzet',index='RokokName_S',columns='TimePeriode',aggfunc={'Omzet':'sum'})
    df12 = df11.T.assign(TotalOmset = lambda x: x.sum(axis=1))
    df13 = df12[['TotalOmset']]
    
    fig = px.line(df13, title='Omset DSO Pamekasan Sem 2 2022')
    fig.update_layout(
    legend_title="Omset DSO",
    xaxis_title="Periode",
    yaxis_title="Ball",
        )

    st.plotly_chart(fig, use_container_width=True)

    #### Omset by Zona #######

    ###@st.cache(allow_output_mutation=True)

    def load_data_8():
        df40 = pd.read_excel(
            io="Omset_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Omset_Sem2_2022",
            usecols="E:I",
            nrows=250000
            )
        return (df40)
    df40=load_data_8()

    df41 = df40.drop(['RokokProductGroup','RokokName_S'], axis=1)
    df42 = df41.pivot_table('Omzet',index='GeographyZona',columns='TimePeriode',aggfunc={'Omzet':'sum'}).T 
    
    fig = px.line(df42, title='Omset By Zona')
    fig.update_layout(
    xaxis_title="Periode",
    yaxis_title="Ball"
        )
    
    st.plotly_chart(fig, use_container_width=True)

    ######## TOP 10 Brand##########

    df20 = dfOmset_Sem2_2022.pivot_table('Omzet',index='RokokName_S',columns='TimePeriode',aggfunc={'Omzet':'sum'}).T
    df21 = df20.T.assign(Rata2 = lambda x: x.mean(axis=1))
    df22 = df21.sort_values('Rata2',ascending = False).groupby('RokokName_S').head(2)
    df23 = df22.head(10).T
    df24 = df23.head(2) ####<<< (periode minggu yg tampil di line chart) ganti tiap minggunya

    fig = px.line(df24, title='Top 10 Brand')
    fig.update_layout(
    legend_title="Brand",
    xaxis_title="Periode",
    yaxis_title="Ball"
        )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("""---""")

    if choose == "DSO Sales Performance":
        st.title(" 💰 Kontribusi Omset")
        ###@st.cache(allow_output_mutation=True)
        def load_data_9():
            df = pd.read_excel(
                io="Market_DSO.xlsx",
                engine="openpyxl",
                sheet_name="MARKET",
                usecols="B:E",
                nrows=1000,
            )
            return (df)
        df=load_data_9()
    
        ###@st.cache(allow_output_mutation=True)
        def load_data_10():
            df = pd.read_excel(
                io="MIT.xlsx",
                engine="openpyxl",
                sheet_name="Sektor",
                usecols="A:P",
                nrows=10000,
            )
            return (df)
        df=load_data_10()
    
        ###@st.cache(allow_output_mutation=True)

        def load_data_11():    
            df1 = pd.read_excel(
                io="Omset_Sem2.xlsx",
                engine="openpyxl",
                sheet_name="Omset_Sem2_2022",
                usecols="A:J",
                nrows=250000,
                )
            return (df1)
        df1=load_data_11()

###        st.write(df)
    
    if choose == "DSO Sales Performance":
        col1, col2 = st.columns(2)
    with col1 :

            fig = px.pie(df1, values='OmzetJtBtg', names='GeographyZona',width= 500, height= 400, title="Kontribusi Omset By Geography Zona Sem 2 2022")
    
            st.plotly_chart(fig)

    with col2 :
            fig = px.pie(df, values='Jun22', names='Kategori-Kelas',width= 500, height= 400, title="Kontribusi Omset By Group Kategori Rokok Sem 2 2022")
    
            st.plotly_chart(fig)


    if choose == "DSO Sales Performance":
        col1, col2 = st.columns(2)
    with col1 :

        fig = px.pie(df1, values='OmzetJtBtg', names='CustomerGroupCategory',width= 500, height= 400, title="Kontribusi Omset By Customer Group Category Sem 2 2022")
        
        st.plotly_chart(fig)
    with col2 :

        fig = px.pie(df1, values='OmzetJtBtg', names='CustomerCategory',width= 500, height= 400, title="Kontribusi Omset By CustomerCategory Sem 2 2022")
        
        st.plotly_chart(fig)  

    ###if option == "DSO Sales Performance":
    ###    st.markdown("""---""")
###
    ###    st.markdown('### Kontribusi Omset Shop By Zona Sem 2 2022')
    ###    fig = plt.figure(figsize=(12, 5.0625))
    ###    ax1 = fig.add_subplot(121)
    ###    ax2 = fig.add_subplot(122)
    ###    fig.subplots_adjust(wspace=0)
###
    ###    # large pie chart parameters
###
    ###    pivot1 = pd.pivot_table(data=df1, index=['CustomerGroupCategory'], values = 'OmzetJtBtg', aggfunc='sum')
    ###    xxx1 = [pivot1.iloc[1]['OmzetJtBtg'], pivot1.iloc[0]['OmzetJtBtg'], pivot1.iloc[2]['OmzetJtBtg'], pivot1.iloc[3]['OmzetJtBtg']]
    ###    pivot2 = pd.pivot_table(data=df1, index=['GeographyZona', 'CustomerGroupCategory'], values = 'OmzetJtBtg', aggfunc='sum')
    ###    yyy1 = [pivot2.iloc[1]['OmzetJtBtg'], pivot2.iloc[5]['OmzetJtBtg'], pivot2.iloc[9]['OmzetJtBtg'], pivot2.iloc[13]['OmzetJtBtg']]
###
    ###    ratios = xxx1
    ###    labels = ['Shop', 'Modern Retail', 'Small Retail', 'Special Retail']
    ###    explode = [0.1, 0, 0, 0]
    ###    # rotate so that first wedge is split by the x-axis
    ###    angle = 60 * ratios[0]
    ###    ax1.pie(ratios, autopct='%1.1f%%', startangle=angle,
    ###            labels=labels, explode=explode)
    ###    # small pie chart parameters
    ###    ratios = yyy1
    ###    labels = ['Bangkalan', 'Pamekasan', 'Sampang', 'Sumenep']
    ###    width = .2
    ###    angle = 60
    ###    ax2.pie(ratios, autopct='%1.1f%%', startangle=angle,
    ###            labels=labels, radius=0.5, textprops={'size': 'smaller'})
    ###    ax1.set_title('Customer Group Category')
    ###    ax2.set_title('Kontribusi Zona')
    ###    # use ConnectionPatch to draw lines between the two plots
    ###    # get the wedge data
    ###    theta1, theta2 = ax1.patches[0].theta1, ax1.patches[0].theta2
    ###    center, r = ax1.patches[0].center, ax1.patches[0].r
    ###    # draw top connecting line
    ###    x = r * np.cos(np.pi / -90 * theta2) + center[0]
    ###    y = np.sin(np.pi / -90 * theta2) + center[1]
    ###    con = ConnectionPatch(xyA=(- width / 2, .5), xyB=(x, y),
    ###                        coordsA="data", coordsB="data", axesA=ax2, axesB=ax1)
    ###    con.set_color([0, 0, 0])
    ###    con.set_linewidth(.5)
    ###    ax2.add_artist(con)
    ###    # draw bottom connecting line
    ###    x = r * np.cos(np.pi / -100 * theta1) + center[0]
    ###    y = np.sin(np.pi / -100 * theta1) + center[1]
    ###    con = ConnectionPatch(xyA=(- width / 2, -.5), xyB=(x, y), coordsA="data",
    ###                        coordsB="data", axesA=ax2, axesB=ax1)
    ###    con.set_color([0, 0, 0])
    ###    ax2.add_artist(con)
    ###    con.set_linewidth(.5)
    ###    st.write(fig)  

############ DSO Canvas Performance #############
  
if choose == "DSO Canvas Performance":

    st.title("💰 Kontribusi Omset Kanvas by Zona Sem 2 2022")

    fig = plt.figure(figsize=(12, 5.0625))
    ax1 = fig.add_subplot(121)
    ax2 = fig.add_subplot(122)
    fig.subplots_adjust(wspace=0)
    
    ###@st.cache(allow_output_mutation=True)

    def load_data_12():    
        df = pd.read_excel(
            io="Omset_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Omset_Sem2_2022",
            usecols="A:J",
            nrows=250000,
            )
        return (df)
    df=load_data_12()

    # large pie chart parameters 

    pivot1 = pd.pivot_table(data=df, index=['CustomerGroupCategory'], values = 'OmzetJtBtg', aggfunc='sum')
    xxx1 = [pivot1.iloc[1]['OmzetJtBtg'], pivot1.iloc[0]['OmzetJtBtg'], pivot1.iloc[2]['OmzetJtBtg'], pivot1.iloc[3]['OmzetJtBtg']]
    pivot2 = pd.pivot_table(data=df, index=['GeographyZona', 'CustomerGroupCategory'], values = 'OmzetJtBtg', aggfunc='sum')
    yyy1 = [pivot2.iloc[2]['OmzetJtBtg'], pivot2.iloc[6]['OmzetJtBtg'], pivot2.iloc[10]['OmzetJtBtg'], pivot2.iloc[14]['OmzetJtBtg']]

    ratios = xxx1
    labels = ['Shop', 'Modern Retail', 'Small Retail', 'Special Retail']
    explode = [0, 0, 0.1, 0]
    # rotate so that first wedge is split by the x-axis
    angle = 59.71 * ratios[0]
    ax1.pie(ratios, autopct='%1.1f%%', startangle=angle,
            labels=labels, explode=explode)
    # small pie chart parameters
    ratios = yyy1
    labels = ['Bangkalan', 'Pamekasan', 'Sampang', 'Sumenep']
    width = .2
    angle = 60
    ax2.pie(ratios, autopct='%1.1f%%', startangle=angle,
            labels=labels, radius=0.5, textprops={'size': 'smaller'})
    ax1.set_title('Customer Group Category')
    ax2.set_title('Geography Zona')
    # use ConnectionPatch to draw lines between the two plots
    # get the wedge data
    theta1, theta2 = ax1.patches[2].theta1, ax1.patches[2].theta2
    center, r = ax1.patches[2].center, ax1.patches[2].r
    # draw top connecting line
    x = r * np.cos(np.pi / 180 * theta2) + center[0]
    y = np.sin(np.pi / 180 * theta2) + center[1]
    con = ConnectionPatch(xyA=(- width / 2, .5), xyB=(x, y),
                        coordsA="data", coordsB="data", axesA=ax2, axesB=ax1)
    con.set_color([0, 0, 0])
    con.set_linewidth(.5)
    ax2.add_artist(con)
    # draw bottom connecting line
    x = r * np.cos(np.pi / 180 * theta1) + center[0]
    y = np.sin(np.pi / 180 * theta1) + center[1]
    con = ConnectionPatch(xyA=(- width / 2, -.5), xyB=(x, y), coordsA="data",
                        coordsB="data", axesA=ax2, axesB=ax1)
    con.set_color([0, 0, 0])
    ax2.add_artist(con)
    con.set_linewidth(.5)
    st.write(fig)  

    st.markdown("""---""")

###    st.markdown('### TDN NIELSEN')
###
###    names = ['LA Bold', 'Geo Mild', 'Chief Filter','L.A. Lights Regular','D. 7 6', 'D. Super','Chief Kretek','D. 7 6 Madu Hitam', 'Gaze Kretek','Ferro Filter']
###
###    df = pd.read_excel(
###        io="/Users/ekasulawestara/Desktop/Dashboard_PMK_2022/st-multi_app/pages/Market_DSO.xlsx",
###        engine="openpyxl",
###        sheet_name="TDN_Nielsen",
###        usecols="A:Q",
###        nrows=300
###        )
###
###    df1 = df.iloc[:,:1].T ###Header Brand
###    df2 = df.drop(['Brand','Group', 'Kategori','Kelas'], axis=1).T ###Body Data
###    df3 = df2.applymap(str)
###
###    df4 = pd.concat([df1, df3], ignore_index = True, axis = True)
###    df5 = df1.append(df3)
###    df6 = df5.iloc[:, 0:10]
###    df6.columns = df6.iloc[0]
###    df6 = df6[1:]
###    df7 = df6.applymap(int)
###    fig = px.line(df7, title='TDN Nielsen Top 10 Brand')
###
###    st.plotly_chart(fig, use_container_width=True)
###
###    df8 = df[df.Brand.isin(names)]
###    df9 = df8.drop(['Group','Kategori','Kelas'], axis=1)
###    df10= df9
###
###
###    df11 = df10.iloc[:,:1].T ###Header Brand
###    df12 = df9.drop(['Brand'], axis=1).T ###Body Data
###    df13 = df12.applymap(str)
###
###    df14 = pd.concat([df11, df13], ignore_index = True, axis = True)
###    df15 = df11.append(df13)
###    df16 = df15.iloc[:, 0:10]
###    df16.columns = df16.iloc[0]
###    df16 = df16[1:]
###    df17 = df16.applymap(int)
###
###    fig = px.line(df17, title='TDN Nielsen Top 10 Brand Djarum')
###
###    st.plotly_chart(fig, use_container_width=True)
###
    st.title(" 🆚 OC vs OV")

    ###@st.cache(allow_output_mutation=True)

    def load_data_13():
        df50 =  pd.read_excel(
                io="S_TrenOVOCKumulatif_data_Sem2.xlsx",
                engine="openpyxl",
                sheet_name="TrenOVOCKumulatif",
                usecols="A:G",
                nrows=7000
                )
        return df50

    df50 = load_data_13()

    df52 = df50.drop(['.View2','Measure Names','Measure Values'], axis=1)

    col1,col2 = st.columns([1,4])
    with col1 :
        Team = st.selectbox("Nama Team",
                                ["Bangkalan Kota",
                                 "Tanah Merah 1",
                                 "Tanah Merah 2",
                                 "Tanjung Bumi",
                                 "Sampang Kota",
                                 "Blega",
                                 "Ketapang",
                                 "Pamekasan Kota",
                                 "Waru",
                                 "Larangan",
                                 "Pademawu",
                                 "Sumenep Kota",
                                 "Ganding",
                                 "Manding"]
                                 )
    with col2 :
        Periode_Minggu = st.multiselect(
            "Pilih Minggu:",
            options=df50["Periode_Minggu"].unique(),
            default=df50["Periode_Minggu"].unique()
            )

    df_selection50 = df50.query("Team==@Team & Periode_Minggu==@Periode_Minggu")
    
    table1 = pd.pivot_table( data=df_selection50, 
                index=['Team','Periode_Minggu'], 
                columns=['Measure Names'], 
                values='.%OC2',
                aggfunc='mean',
                fill_value=0
                        )

    table2 = pd.pivot_table( data=df_selection50, 
            index=['Team','Periode_Minggu'], 
            columns=['Measure Names'], 
            values='.%OV2',
            aggfunc='mean',
            fill_value=0
                    )

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=Periode_Minggu,
        y=table1['.%OC2']*100,
        text= (round(table1['.%OC2'],2)*100),
        name='OC%',
        marker_color='lightsalmon'
    ))
    fig.add_trace(go.Bar(
        x=Periode_Minggu,
        y=table2['.%OV2']*100,
        text= (round(table2['.%OV2'],2)*100),
        name='OV%',
        marker_color='indianred'
    ))
    fig.update_layout(title= Team, barmode='group', autosize=True,xaxis_title='Minggu', yaxis_title='%')
    fig.update_traces(textposition='outside')
    st.plotly_chart(fig, use_container_width=True)
###
########## KPI Realisasi Kanvas-TDN-EC ###########

if choose == "KPI Realisasi Kanvas-TDN-EC":

################# D.Super ################

    ###@st.cache(allow_output_mutation=True)

    def load_data_14():
        df = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Pivot_Super",
            usecols="X:AS",
            nrows=52
            )
        return df

    df = load_data_14()

    st.title(":chart_with_upwards_trend: KPI Promotor D.Super 12")

    col1,col2 = st.columns([1,4])
    with col1 :
        Nama_Promotor = st.selectbox("",["05004102-Rofi Wijayanto", "05004280-Etza Enggar Adikara", "05004525-Baha 'uddin","05006202-Ibnu Hajar", "05004941-Hermansyah Budi Prasetia", "05005803-Fany Firmansyah", "05007039-Mohammad Hadi Wahyudi", "05006341-Nurcholis Alamy", "05006923-Bareza Fuadi Almaufiri", "05006394-Yudhi Imam Saputra", "05003210-Ahmad Santoso SV", "05003752-Fahrul Antoni", "05006074-Fathor Arifin"])
        
    with col2 :
        Minggu = st.multiselect(
        "Pilih Minggu:",
        options=df["Minggu"].unique(),
        default=df["Minggu"].unique()
        )
    st.markdown("##")

    df_selection = df.query("Minggu == @Minggu & Nama_Promotor ==@Nama_Promotor")
    Realisasi_Canvas_Plan = (round(df_selection["Realisasi_Canvas_Plan"].mean(),2)*100)
    TDN_Sebelum = (round(df_selection["TDN_Sebelum"].mean(), 2)*100)
    EC = (round(df_selection["EC"].mean(), 2)*100)

    Nama_Promotor = (df_selection["Nama_Promotor"])
    Minggu = (df_selection["Minggu"])


###    st.header("D.SUPER 12")

    left_column, middle_column, right_column = st.columns(3)
    with left_column: 
            task_Realisasi_Kanvas_Plan =st.subheader("Averg. Realisasi Kanvas :") 
            task_Realisasi_Kanvas_Plan =st.subheader(f"{Realisasi_Canvas_Plan} %")
            

    with middle_column:
            task_TDN_Sebelum =st.subheader("Averg.TDN Sebelum :") 
            task_TDN_Sebelum =st.subheader(f"{TDN_Sebelum} %")
                     
            
    with right_column:
            task_EC =st.subheader("Averg. EC D.Super 12:") 
            task_EC =st.subheader(f"{EC} %")
            

    st.markdown("""---""")
#####



    # ---- MAINPAGE ----



    Nama_Promotor = (df_selection["Nama_Promotor"])
    Minggu = (df_selection["Minggu"])
    ###@st.cache(allow_output_mutation=True)

    def load_data_15():
        df = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Pivot_Super",
            usecols="X:AS",
            nrows=575,
        )
        return df

    df = load_data_15()

    fig = plt.figure()
    
    fig.set_figheight(2)
    fig.set_figwidth(10)

    st.subheader('Realisasi Kanvas vs TDN Sebelum vs EC D.Super 12')
    plt.plot(df_selection['Minggu'],df_selection['EC'], color='blue', marker='o')
    plt.plot(df_selection['Minggu'],df_selection['TDN_Sebelum'], color='red', marker='o')
    plt.plot(df_selection['Minggu'],df_selection['Realisasi_Canvas_Plan'], color='green', marker='o')
    plt.legend(['EC', 'TDN Sebelum', 'Realisasi Kanvas'], bbox_to_anchor=(0.95, 1.0, 0.3, 0.2),loc ='upper right')
    plt.xlabel('Minggu Ke-', fontsize=10)
    plt.xticks(rotation=45, ha='right')
    plt.ylabel('x 100%', fontsize=10)

    plt.grid(True)

    
    st.set_option('deprecation.showPyplotGlobalUse', False)

    st.pyplot()

    with st.expander ("Show Detail"):

                    st.write("Data Realisasi Kanvas Plan")

                    table = pd.pivot_table( data=df_selection, 
                                    index=['Nama_Promotor'], 
                                    columns=['Minggu'], 
                                    values='Realisasi_Canvas_Plan',
                                    aggfunc='mean',
                                    fill_value=0
                                            )

                    st.dataframe(table)

                    st.write("Data TDN Sebelum Kanvas D.Super 12")

                    table = pd.pivot_table( data=df_selection, 
                                    index=['Nama_Promotor'], 
                                    columns=['Minggu'], 
                                    values="TDN_Sebelum",
                                    aggfunc='mean',
                                    fill_value=0
                                            )
                    st.dataframe(table)


                    st.write("Data EC D.Super 12")

                    table = pd.pivot_table( data=df_selection, 
                                    index=['Nama_Promotor'], 
                                    columns=['Minggu'], 
                                    values='EC',
                                    aggfunc='mean',
                                    fill_value=0
                                            )
                    st.dataframe(table)

################# 76 Madu ################

    ###@st.cache(allow_output_mutation=True)

    def load_data_16():
        df = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Pivot_76Madu",
            usecols="X:AS",
            nrows=575
        )
        return df

    df = load_data_16()
    
    st.markdown("""___""")

    st.title(":chart_with_upwards_trend: KPI Promotor D.76 Madu Hitam")

    col1,col2 = st.columns([1,4])
    with col1 :
        Nama_Promotor1 = st.selectbox("",["05004102-Rofi Wijayanto.", "05004280-Etza Enggar Adikara.", "05004525-Baha 'uddin.","05006202-Ibnu Hajar.", "05004941-Hermansyah Budi Prasetia.", "05005803-Fany Firmansyah.", "05007039-Mohammad Hadi Wahyudi.", "05006341-Nurcholis Alamy.", "05006923-Bareza Fuadi Almaufiri.", "05006394-Yudhi Imam Saputra.", "05003210-Ahmad Santoso SV.", "05003752-Fahrul Antoni.", "05006074-Fathor Arifin."])
    
    with col2 :
        Minggu_ = st.multiselect(
        "Pilih Minggu:",
        options=df["Minggu_"].unique(),
        default=df["Minggu_"].unique()
        )
    st.markdown("##")

    df_selection = df.query("Minggu_==@Minggu_ & Nama_Promotor1==@Nama_Promotor1")

    Realisasi_Canvas_Plan = (round(df_selection["Realisasi_Canvas_Plan"].mean(),2)*100)
    TDN_Sebelum = (round(df_selection["TDN_Sebelum"].mean(), 2)*100)
    EC = (round(df_selection["EC"].mean(), 2)*100)


    Nama_Promotor1 = (df_selection["Nama_Promotor1"])
    Minggu1 = (df_selection["Minggu_"])


###    st.header("D.76 Madu Hitam 12")

    left_column, middle_column, right_column = st.columns(3)
    with left_column: 
            task_Realisasi_Kanvas_Plan =st.subheader("Averg. Realisasi Kanvas :") 
            task_Realisasi_Kanvas_Plan =st.subheader(f"{Realisasi_Canvas_Plan} %")
            

    with middle_column:
            task_TDN_Sebelum =st.subheader("Averg.TDN Sebelum :") 
            task_TDN_Sebelum =st.subheader(f"{TDN_Sebelum} %")
                     
            
    with right_column:
            task_EC =st.subheader("Averg. EC D.76 Madu H 12:") 
            task_EC =st.subheader(f"{EC} %")
            

    st.markdown("""---""")


    # ---- MAINPAGE ----



    Nama_Promotor1 = (df_selection["Nama_Promotor1"])
    Minggu1 = (df_selection["Minggu_"])

    ###@st.cache(allow_output_mutation=True)

    def load_data_17():
        df = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Pivot_76Madu",
            usecols="X:BD",
            nrows=1000,
        )
        return df


    df = load_data_17()

    fig = plt.figure()
    
    fig.set_figheight(2)
    fig.set_figwidth(10)

    st.subheader('Realisasi Kanvas vs TDN Sebelum vs EC D.76 Madu H 12')
    
    plt.plot(df_selection['Minggu_'],df_selection['EC'], color='blue', marker='o')
    plt.plot(df_selection['Minggu_'],df_selection['TDN_Sebelum'], color='red', marker='o')
    plt.plot(df_selection['Minggu_'],df_selection['Realisasi_Canvas_Plan'], color='green', marker='o')
    plt.legend(['EC', 'TDN Sebelum', 'Realisasi Kanvas'], bbox_to_anchor=(0.95, 1.0, 0.3, 0.2),loc ='upper right')
    plt.xlabel('Minggu Ke-', fontsize=10)
    plt.xticks(rotation=45, ha='right')
    plt.ylabel('x 100%', fontsize=10)

    plt.grid(True)

    st.set_option('deprecation.showPyplotGlobalUse', False)

    st.pyplot() 

    with st.expander ("Show Detail"):

                    st.write("Data Realisasi Kanvas Plan")

                    table = pd.pivot_table( data=df_selection, 
                                    index=['Nama_Promotor1'], 
                                    columns=['Minggu_'], 
                                    values='Realisasi_Canvas_Plan',
                                    aggfunc='mean',
                                    fill_value=0
                                            )

                    st.dataframe(table)

                    st.write("Data TDN Sebelum Kanvas D.76 Madu H 12")

                    table = pd.pivot_table( data=df_selection, 
                                    index=['Nama_Promotor1'], 
                                    columns=['Minggu_'], 
                                    values="TDN_Sebelum",
                                    aggfunc='mean',
                                    fill_value=0
                                            )
                    st.dataframe(table)


                    st.write("Data EC D.76 Madu H 12")

                    table = pd.pivot_table( data=df_selection, 
                                    index=['Nama_Promotor1'], 
                                    columns=['Minggu_'], 
                                    values='EC',
                                    aggfunc='mean',
                                    fill_value=0
                                            )
                    st.dataframe(table)



############# Rank Promotor #############

if choose == "Rank Promotor":

    st.title(" 🏆 Rank Promotor")

        ####################### RANK KPI ALL PROMOTOR #######################

    ############################ KPI D.Super ############################

    ###@st.cache(allow_output_mutation=True)

    def load_data_18():
        df1 = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Super",
            usecols="A:E",
            nrows=15000
        )
        return df1
    df1 = load_data_18()

    ############# Poin & Target KPI ##############

    target_EC_Super = 0.3      ### <<<< ubah target disini
    target_TDN_Super = 0.8     ### <<<< ubah target disini 
    target_Kanvas_Plan = 0.85  ### <<<< ubah target disini
    target_EC_76Madu = 0.3     ### <<<< ubah target disini
    target_TDN_76Madu = 0.8    ### <<<< ubah target disini 

    poin_ov = 0.15             ### <<<< ubah poin disini
    poin_oc = 0.15             ### <<<< ubah poin disini
    poin_kanvas_plan = 0.15    ### <<<< ubah poin disini
    poin_ec_super = 0.15       ### <<<< ubah poin disini
    poin_ec_76madu = 0.1       ### <<<< ubah poin disini
    poin_tdn_super = 0.15      ### <<<< ubah poin disini
    poin_tdn_76madu = 0.1      ### <<<< ubah poin disini

    #############################################

    data_avg_rank_promotor_super =pd.pivot_table( data=df1, 
                            index=['Nama_Promotor'], 
                            columns=['Measure'], 
                            values='Values',
                            aggfunc='mean',
                            fill_value=0
                                    )

    data_ec_avg_rank_promotor_super = data_avg_rank_promotor_super.loc[:,["EC"]]
    data_tdn_avg_rank_promotor_super = data_avg_rank_promotor_super.loc[:,["TDN_Sebelum"]]
    data_kanvas_plan_avg_rank_promotor_super = data_avg_rank_promotor_super.loc[:,["Realisasi_Canvas_Plan"]]

    data_ec_avg_rank_promotor_super['Target_EC_Super'] = data_ec_avg_rank_promotor_super['EC'].apply(lambda x: target_EC_Super if x >= target_EC_Super else (target_EC_Super))
    bobot_ec_super = data_ec_avg_rank_promotor_super

    data_tdn_avg_rank_promotor_super['Target_TDN_Super'] = data_tdn_avg_rank_promotor_super['TDN_Sebelum'].apply(lambda x: target_TDN_Super if x >= target_TDN_Super else (target_TDN_Super))
    bobot_tdn_super = data_tdn_avg_rank_promotor_super

    data_kanvas_plan_avg_rank_promotor_super['Target_Kanvas_Plan'] = data_kanvas_plan_avg_rank_promotor_super['Realisasi_Canvas_Plan'].apply(lambda x: target_Kanvas_Plan if x >= target_Kanvas_Plan else (target_Kanvas_Plan))
    bobot_kanvas_plan = data_kanvas_plan_avg_rank_promotor_super
    

    bobot_ec_super['Bobot_EC_Super'] = bobot_ec_super['EC']/bobot_ec_super['Target_EC_Super']
    bobot_ec_super['Nilai_EC_Super'] = bobot_ec_super['Bobot_EC_Super'] *poin_ec_super

    bobot_ec_super['Nilai_EC_Super'] = target_EC_Super/target_EC_Super*poin_ec_super
    bobot_ec_super.loc[bobot_ec_super['Target_EC_Super'] * bobot_ec_super['Bobot_EC_Super'] < target_EC_Super, 'Nilai_EC_Super'] = bobot_ec_super['Bobot_EC_Super'] *poin_ec_super

    bobot_tdn_super['Bobot_TDN_Super'] = bobot_tdn_super['TDN_Sebelum']/bobot_tdn_super['Target_TDN_Super']
    bobot_tdn_super['Nilai_TDN_Super'] = bobot_tdn_super['Bobot_TDN_Super'] *poin_tdn_super  

    bobot_tdn_super['Nilai_TDN_Super'] = target_TDN_Super/target_TDN_Super*poin_tdn_super
    bobot_tdn_super.loc[bobot_tdn_super['Target_TDN_Super'] * bobot_tdn_super['Bobot_TDN_Super'] < target_TDN_Super, 'Nilai_TDN_Super'] = bobot_tdn_super['Bobot_TDN_Super'] * bobot_tdn_super['Target_TDN_Super'] *poin_tdn_super
    
    bobot_kanvas_plan['Bobot_Kanvas_Plan'] = bobot_kanvas_plan['Realisasi_Canvas_Plan']/bobot_kanvas_plan['Target_Kanvas_Plan']
    bobot_kanvas_plan['Nilai_Kanvas_Plan'] = bobot_kanvas_plan['Bobot_Kanvas_Plan'] *poin_kanvas_plan  

    bobot_kanvas_plan['Nilai_Kanvas_Plan'] = target_Kanvas_Plan/target_Kanvas_Plan*poin_kanvas_plan
    bobot_kanvas_plan.loc[bobot_kanvas_plan['Target_Kanvas_Plan'] * bobot_kanvas_plan['Bobot_Kanvas_Plan'] < target_Kanvas_Plan, 'Nilai_Kanvas_Plan'] = bobot_kanvas_plan['Bobot_Kanvas_Plan'] * bobot_kanvas_plan['Target_Kanvas_Plan'] *poin_kanvas_plan
    


    ############################ KPI D.76 Madu ############################
    
    ###@st.cache(allow_output_mutation=True)

    def load_data_19():
        df2 = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="76Madu",
            usecols="A:E",
            nrows=15000
        )
        return df2
    df2 = load_data_19()

    data_avg_rank_promotor_76Madu =pd.pivot_table( data=df2, 
                            index=['Nama_Promotor'], 
                            columns=['Measure'], 
                            values='Values',
                            aggfunc='mean',
                            fill_value=0
                                    )

    data_ec_avg_rank_promotor_76Madu = data_avg_rank_promotor_76Madu.loc[:,["EC"]]
    data_tdn_avg_rank_promotor_76Madu = data_avg_rank_promotor_76Madu.loc[:,["TDN_Sebelum"]]

    data_ec_avg_rank_promotor_76Madu['Target_EC_76Madu'] = data_ec_avg_rank_promotor_76Madu['EC'].apply(lambda x: target_EC_76Madu if x >= target_EC_76Madu else (target_EC_76Madu))
    bobot_ec_76Madu = data_ec_avg_rank_promotor_76Madu
    data_tdn_avg_rank_promotor_76Madu['Target_TDN_76Madu'] = data_tdn_avg_rank_promotor_76Madu['TDN_Sebelum'].apply(lambda x: target_TDN_76Madu if x >= target_TDN_76Madu else (target_TDN_76Madu))
    bobot_tdn_76Madu = data_tdn_avg_rank_promotor_76Madu
    

    bobot_ec_76Madu['Bobot_EC_76Madu'] = bobot_ec_76Madu['EC']/bobot_ec_76Madu['Target_EC_76Madu']
    bobot_ec_76Madu['Nilai_EC_76Madu'] = bobot_ec_76Madu['Bobot_EC_76Madu'] *poin_ec_76madu

    bobot_ec_76Madu['Nilai_EC_76Madu'] = target_EC_76Madu/target_EC_76Madu*poin_ec_76madu
    bobot_ec_76Madu.loc[bobot_ec_76Madu['Target_EC_76Madu'] * bobot_ec_76Madu['Bobot_EC_76Madu'] < target_EC_76Madu, 'Nilai_EC_76Madu'] = bobot_ec_76Madu['Bobot_EC_76Madu'] *poin_ec_76madu

    bobot_tdn_76Madu['Bobot_TDN_76Madu'] = bobot_tdn_76Madu['TDN_Sebelum']/bobot_tdn_76Madu['Target_TDN_76Madu']
    bobot_tdn_76Madu['Nilai_TDN_76Madu'] = bobot_tdn_76Madu['Bobot_TDN_76Madu'] *poin_tdn_76madu  
    

    bobot_tdn_76Madu['Nilai_TDN_76Madu'] = target_TDN_76Madu/target_TDN_76Madu*poin_tdn_76madu
    bobot_tdn_76Madu.loc[bobot_tdn_76Madu['Target_TDN_76Madu'] * bobot_tdn_76Madu['Bobot_TDN_76Madu'] < target_TDN_76Madu, 'Nilai_TDN_76Madu'] = bobot_tdn_76Madu['Bobot_TDN_76Madu'] * bobot_tdn_76Madu['Target_TDN_76Madu'] *poin_tdn_76madu

##################################################
    ###@st.cache(allow_output_mutation=True)

    def load_data_20():
        df_ovoc =  pd.read_excel(
                    io="S_TrenOVOCKumulatif_data_Sem2.xlsx",
                    engine="openpyxl",
                    sheet_name="KPIOVOC",
                    usecols="A:G",
                    nrows=7000
                    )
        return df_ovoc

    df_ovoc = load_data_20()

    table_oc1 = pd.pivot_table( data=df_ovoc, 
                index=['Team'], 
                columns=['Measure Names'], 
                values='OC2',
                aggfunc='mean',
                fill_value=0
                        )
    table_ov2 = pd.pivot_table( data=df_ovoc, 
            index=['Team'], 
            columns=['Measure Names'], 
            values='OV2',
            aggfunc='mean',
            fill_value=0
                    )
    table_oc1_1 = table_oc1.loc[:,'OC2']
    table_ov2_1 = table_ov2.loc[:,'OV2']
    table_oc1['Target_OC'] = table_oc1['OC2'].apply(lambda x: 1 if x >= 1 else (1))
    table_ov2['Target_OV'] = table_ov2['OV2'].apply(lambda x: 1 if x >= 1 else (1))
    table_oc1['Bobot_OC'] = table_oc1['OC2']/table_oc1['Target_OC']
    table_oc1['Nilai_OC'] = table_oc1['Bobot_OC'] *poin_oc
    table_ov2['Bobot_OV'] = table_ov2['OV2']/table_ov2['Target_OV']
    table_ov2['Nilai_OV'] = table_ov2['Bobot_OV'] *poin_ov 
    table_oc1['Nilai_OC'] = 1/1*0.15
    table_oc1.loc[table_oc1['Target_OC'] * table_oc1['Bobot_OC'] < 1, 'Nilai_OC'] = table_oc1['Bobot_OC'] *poin_oc
    table_ov2['Nilai_OV'] = 1/1*0.15
    table_ov2.loc[table_ov2['Target_OV'] * table_ov2['Bobot_OV'] < 1, 'Nilai_OV'] = table_ov2['Bobot_OV'] * table_ov2['Target_OV'] *poin_ov
    table_oc1_1 = table_oc1.drop(['OV2'], axis=1)
    table_ov2_1 = table_ov2.drop(['OC2'], axis=1)
    bobot_kanvas_plan_final = bobot_kanvas_plan.loc[:,['Nilai_Kanvas_Plan']]
    bobot_tdn_super_final = bobot_tdn_super.loc[:,['Nilai_TDN_Super']]
    bobot_ec_super_final = bobot_ec_super.loc[:,['Nilai_EC_Super']]
    bobot_tdn_76Madu_final = bobot_tdn_76Madu.loc[:,['Nilai_TDN_76Madu']]
    bobot_ec_76Madu_final = bobot_ec_76Madu.loc[:,['Nilai_EC_76Madu']]
    table_oc1_1_final = table_oc1_1.loc[:,['Nilai_OC']]
    table_ov2_1_final = table_ov2_1.loc[:,['Nilai_OV']]
    evaluasi = pd.concat([bobot_kanvas_plan_final, bobot_tdn_super_final,bobot_ec_super_final,bobot_tdn_76Madu_final,bobot_ec_76Madu_final,table_oc1_1_final,table_ov2_1_final], axis = True)
    evaluasi['Total Nilai'] = (evaluasi.sum(axis=1))
    ###evaluasi = evaluasi.loc[:,['Total Nilai']].sort_values(by="Total Nilai",ascending=False)
    st.write(evaluasi)
    with st.expander("Show Data"):
        st.write(bobot_kanvas_plan)
        st.write(bobot_ec_super)
        st.write(bobot_tdn_super)
        st.write(bobot_ec_76Madu)
        st.write(bobot_tdn_76Madu)
        st.write(table_oc1_1)
        st.write(table_ov2_1)

############### D.Super ################

    ###@st.cache(allow_output_mutation=True)

    def load_data_21():
        df = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Pivot_Super2",
            usecols="AD:BE",
            nrows=500
        )
        return df

    df = load_data_21()


#####################################################



    st.title("🥇🥈🥉 KPI by Promotor - D.Super 12")

    col1,col2 = st.columns([1,4])
    with col1 :
        Nama_Measure = st.selectbox("",["EC", "TDN_Sebelum", "Realisasi_Canvas_Plan"])
        
    with col2 :
        Minggu = st.multiselect(
        "Pilih Minggu:",
        options=df["Minggu"].unique(),
        default=df["Minggu"].unique()
        )
    st.markdown("##")

    df_selection = df.query("Minggu == @Minggu & Measure ==@Nama_Measure")
    Nama_Promotor1 = (df_selection["05003210-Ahmad Santoso SV"])
    Nama_Promotor2 = (df_selection["05003752-Fahrul Antoni"])
    ###Nama_Promotor3 = (df_selection["05003888-Ach Safi'i"])
    Nama_Promotor4 = (df_selection["05004102-Rofi Wijayanto"])
    Nama_Promotor5 = (df_selection["05004280-Etza Enggar Adikara"])
    Nama_Promotor6 = (df_selection["05004525-Baha 'uddin"])
    Nama_Promotor7 = (df_selection["05006923-Bareza Fuadi Almaufiri"])
    ###Nama_Promotor8 = (df_selection["05004855-Jaka Ade Saputra"])
    Nama_Promotor9 = (df_selection["05004941-Hermansyah Budi Prasetia"])
    Nama_Promotor10 = (df_selection["05005803-Fany Firmansyah"])
    Nama_Promotor11 = (df_selection["05006074-Fathor Arifin"])
    Nama_Promotor12 = (df_selection["05006202-Ibnu Hajar"])
    Nama_Promotor13 = (df_selection["05007039-Mohammad Hadi Wahyudi"])
    Nama_Promotor14 = (df_selection["05006341-Nurcholis Alamy"])
    ###Nama_Promotor15 = (df_selection["05006342-Achmad Junaidi Saleh"])
    Nama_Promotor16 = (df_selection["05006394-Yudhi Imam Saputra"])

    st.header("D.SUPER 12")


    #### ---- MAINPAGE ----



    Nama_Measure = (df_selection["Measure"])
    Minggu = (df_selection["Minggu"])

    ###@st.cache(allow_output_mutation=True)

    def load_data_22():
        df = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Pivot_Super2",
            usecols="AD:BE",
            nrows=500,
        )
        return df

    df = load_data_22()


    fig = plt.figure()
    
    fig.set_figheight(4)
    fig.set_figwidth(10)

    plt.plot(df_selection['Minggu'], df_selection["05003210-Ahmad Santoso SV"], color='blue', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05003752-Fahrul Antoni"], color='deepskyblue', marker='.')
    ###plt.plot(df_selection['Minggu'], df_selection["05003888-Ach Safi'i"], color='red', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05004102-Rofi Wijayanto"], color='sienna', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05004280-Etza Enggar Adikara"], color='deeppink', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05004525-Baha 'uddin"], color='indigo', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05006923-Bareza Fuadi Almaufiri"], color='tomato', marker='.')
    ###plt.plot(df_selection['Minggu'], df_selection["05004855-Jaka Ade Saputra"], color='gold', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05004941-Hermansyah Budi Prasetia"], color='dimgray', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05005803-Fany Firmansyah"], color='darkkhaki', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05006074-Fathor Arifin"], color='crimson', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05006202-Ibnu Hajar"], color='teal', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05007039-Mohammad Hadi Wahyudi"], color='black', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05006341-Nurcholis Alamy"], color='olivedrab', marker='.')
    ###plt.plot(df_selection['Minggu'], df_selection["05006342-Achmad Junaidi Saleh"], color='fuchsia', marker='.')
    plt.plot(df_selection['Minggu'], df_selection["05006394-Yudhi Imam Saputra"], color='olive', marker='.')
    plt.legend(["05003210-Ahmad Santoso SV",
                "05003752-Fahrul Antoni",
                ###"05003888-Ach Safi'i",
                "05004102-Rofi Wijayanto",
                "05004280-Etza Enggar Adikara",
                "05004525-Baha 'uddin",
                "05006923-Bareza Fuadi Almaufiri",
                "05004941-Hermansyah Budi Prasetia",
                "05005803-Fany Firmansyah",
                "05006074-Fathor Arifin",
                "05006202-Ibnu Hajar",
                "05007039-Mohammad Hadi Wahyudi",
                "05006341-Nurcholis Alamy",
                "05006394-Yudhi Imam Saputra"
                ],
                bbox_to_anchor=(1, 1.0, 0.5, 0.1),loc ='upper right')
    plt.xlabel('Minggu Ke-', fontsize=10)
    plt.xticks(rotation=45, ha='right')
    plt.ylabel('x 100%', fontsize=10)

    plt.grid(True)

    
    st.set_option('deprecation.showPyplotGlobalUse', False)

    st.pyplot()

############### 76 Madu ##############

    ###@st.cache(allow_output_mutation=True)

    def load_data_23():
        df = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Pivot_76Madu2",
            usecols="AD:BE",
            nrows=500
        )
        return df

    df = load_data_23()

    
    st.title("🥇🥈🥉 KPI by Promotor - D.76 Madu 12")


    col1,col2 = st.columns([1,4])
    with col1 :
        Nama_Measure2 = st.selectbox("",["EC.", "TDN_Sebelum.", "Realisasi_Canvas_Plan."])
        
    with col2 :
        Minggu__ = st.multiselect(
        "Pilih Minggu:",
        options=df["Minggu__"].unique(),
        default=df["Minggu__"].unique()
        )

    st.markdown("##")

    df_selection = df.query("Minggu__ == @Minggu__ & Measure2 ==@Nama_Measure2")
    Nama_Promotor1 = (df_selection["05003210-Ahmad Santoso SV"])
    Nama_Promotor2 = (df_selection["05003752-Fahrul Antoni"])
    ###Nama_Promotor3 = (df_selection["05003888-Ach Safi'i"])
    Nama_Promotor4 = (df_selection["05004102-Rofi Wijayanto"])
    Nama_Promotor5 = (df_selection["05004280-Etza Enggar Adikara"])
    Nama_Promotor6 = (df_selection["05004525-Baha 'uddin"])
    Nama_Promotor7 = (df_selection["05006923-Bareza Fuadi Almaufiri"])
    ###Nama_Promotor8 = (df_selection["05004855-Jaka Ade Saputra"])
    Nama_Promotor9 = (df_selection["05004941-Hermansyah Budi Prasetia"])
    Nama_Promotor10 = (df_selection["05005803-Fany Firmansyah"])
    Nama_Promotor11 = (df_selection["05006074-Fathor Arifin"])
    Nama_Promotor12 = (df_selection["05006202-Ibnu Hajar"])
    Nama_Promotor13 = (df_selection["05007039-Mohammad Hadi Wahyudi"])
    Nama_Promotor14 = (df_selection["05006341-Nurcholis Alamy"])
    ###Nama_Promotor15 = (df_selection["05006342-Achmad Junaidi Saleh"])
    Nama_Promotor16 = (df_selection["05006394-Yudhi Imam Saputra"])

    st.header("D.76 MADU 12")


    #### ---- MAINPAGE ----



    Nama_Measure = (df_selection["Measure2"])
    Minggu2 = (df_selection["Minggu__"])

    ###@st.cache(allow_output_mutation=True)

    def load_data_24():
        df = pd.read_excel(
            io="Rekap_Kanvas_Sem2.xlsx",
            engine="openpyxl",
            sheet_name="Pivot_76Madu2",
            usecols="AD:BE",
            nrows=500,
        )
        # Add 'hour' column to dataframe
        return df

    df = load_data_24()

    fig = plt.figure()
    
    fig.set_figheight(4)
    fig.set_figwidth(10)

    plt.plot(df_selection['Minggu__'], df_selection["05003210-Ahmad Santoso SV"], color='blue', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05003752-Fahrul Antoni"], color='deepskyblue', marker='.')
    ###plt.plot(df_selection['Minggu__'], df_selection["05003888-Ach Safi'i"], color='red', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05004102-Rofi Wijayanto"], color='sienna', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05004280-Etza Enggar Adikara"], color='deeppink', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05004525-Baha 'uddin"], color='indigo', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05006923-Bareza Fuadi Almaufiri"], color='tomato', marker='.')
    ###plt.plot(df_selection['Minggu__'], df_selection["05004855-Jaka Ade Saputra"], color='gold', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05004941-Hermansyah Budi Prasetia"], color='dimgray', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05005803-Fany Firmansyah"], color='darkkhaki', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05006074-Fathor Arifin"], color='crimson', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05006202-Ibnu Hajar"], color='teal', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05007039-Mohammad Hadi Wahyudi"], color='black', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05006341-Nurcholis Alamy"], color='olivedrab', marker='.')
    ###plt.plot(df_selection['Minggu__'], df_selection["05006342-Achmad Junaidi Saleh"], color='fuchsia', marker='.')
    plt.plot(df_selection['Minggu__'], df_selection["05006394-Yudhi Imam Saputra"], color='olive', marker='.')
    plt.legend(["05003210-Ahmad Santoso SV",
                "05003752-Fahrul Antoni",
                ###"05003888-Ach Safi'i",
                "05004102-Rofi Wijayanto",
                "05004280-Etza Enggar Adikara",
                "05004525-Baha 'uddin",
                "05006923-Bareza Fuadi Almaufiri",
                "05004941-Hermansyah Budi Prasetia",
                "05005803-Fany Firmansyah",
                "05006074-Fathor Arifin",
                "05006202-Ibnu Hajar",
                "05007039-Mohammad Hadi Wahyudi",
                "05006341-Nurcholis Alamy",
                "05006394-Yudhi Imam Saputra"
                ],
                bbox_to_anchor=(1, 1.0, 0.5, 0.1),loc ='upper right')
    plt.xlabel('Minggu Ke-', fontsize=10)
    plt.xticks(rotation=45, ha='right')
    plt.ylabel('x 100%', fontsize=10)

    plt.grid(True)

    
    st.set_option('deprecation.showPyplotGlobalUse', False)

    st.pyplot()


####file_name_evaluasi = 'Rank_Promotor.xlsx'
####
####
####with pd.ExcelWriter(file_name_evaluasi, engine='xlsxwriter')as writer:
####        evaluasi.to_excel(writer, sheet_name ='Evaluasi')
####        bobot_kanvas_plan.to_excel(writer, sheet_name ='bobot_kanvas_plan')
####        bobot_ec_super.to_excel(writer, sheet_name ='bobot_ec_super')
####        bobot_tdn_super.to_excel(writer, sheet_name ='bobot_tdn_super')
####        bobot_ec_76Madu.to_excel(writer, sheet_name ='bobot_ec_76Madu')
####       bobot_tdn_76Madu.to_excel(writer, sheet_name ='bobot_tdn_76Madu')
####        table_oc1_1.to_excel(writer, sheet_name ='table_oc1_1')
####        table_ov2_1.to_excel(writer, sheet_name ='table_ov2_1')



