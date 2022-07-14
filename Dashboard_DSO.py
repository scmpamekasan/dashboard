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
from matplotlib.patches import ConnectionPatch


st.set_page_config(layout="wide")

####################### HEADER ####################
st.markdown("<h1 style='text-align: center; color: black;'>DASHBOARD DSO PAMEKASAN</h1>", unsafe_allow_html=True)

#################### Menu SideBar ##################


option = st.sidebar.selectbox("Menu",[
                              "Home Page",
                              "Struktur Organisasi",
                              "MP By Zona - Kategori - Kelas",
                              "MS By Group Pabrikan",
                              "Top 10 Brand Nielsen",
                              "Top 10 Brand Internal",
                              "DSO Sales Performance",
                              "DSO Canvas Performance"
                              ]
                              )

####################### MAP DSO ####################

if option == "Home Page":
    st.markdown("""---""")

    height = 200
    width = 1200
    
    df = pd.DataFrame(
        np.random.randn(1, 2) / [50, 50] + [(-7.004320, 113.861274), (-7.187682, 113.240499),(-7.160772, 113.482642), (-7.048228, 112.781668)],
        columns=['lat', 'lon'],)

    st.pydeck_chart(pdk.Deck(
        map_style='mapbox://styles/mapbox/light-v9',
        initial_view_state=pdk.ViewState(
            latitude=-7.060772,
            longitude=113.340499,
            zoom=8,
            height=height,
            width=width
        ),
        layers=[
            pdk.Layer(
                'ScatterplotLayer',
                data=df,
                get_position='[lon, lat]',
                get_color='[200, 30, 0, 160]',
                get_radius=200,
                auto_highlight=True
            ),
        ],
    ))

    ####################### MAIN PAGE ####################

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("MP Jt Batang/Hari", "12,4")
    col2.metric("OU", "48.560")
    col3.metric("Total Penduduk", "4.004.563")
    col4.metric("Jumlah Zona", "4")


    col1, col2, col3, col4 = st.columns(4)
    col1.metric("MP Ball Standard/Mgg", "30.880")
    col2.metric("Outlet Database", "9.413")
    col3.metric("Laki Laki (Smokers)", "1.212.375")
    col4.metric("Jumlah Sektor", "72")

    st.markdown("""---""")
    Zona = st.radio(
        "Select Zona",
        ('Bangkalan', 'Sampang', 'Pamekasan', 'Sumenep'), horizontal= True)

    if Zona == 'Bangkalan':
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("MP Jt Batang/Hari", "3,3")
        col1.metric("MP Ball Standard/Mgg", "8.136")
        col2.metric("OU", "12.541")
        col2.metric("Outlet Database", "2.580")
        col3.metric("Total Penduduk", "1.034.197")
        col3.metric("Laki Laki (Smokers)", "319.431")
        col4.metric("Jumlah Sektor", "17")
        col4.metric("Perkapita Rokok/Mgg 2021 (Rp)", "18.994")

    if Zona == 'Sampang':
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("MP Jt Batang/hari", "3,1")
        col1.metric("MP Ball Standard/Mgg", "7.871")
        col2.metric("OU", "12.076")
        col2.metric("Outlet Database", "2.100")
        col3.metric("Total Penduduk", "995.874")
        col3.metric("Laki Laki (Smokers)", "309.025")
        col4.metric("Jumlah Sektor", "15")
        col4.metric("Perkapita Rokok/Mgg (Rp)", "17.521")

    if Zona == 'Pamekasan':
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("MP Jt Batang/Hari", "2,6")
        col1.metric("MP Ball Standard/Mgg", "6.506")
        col2.metric("OU", "10.308")
        col2.metric("Outlet Database", "2.240")
        col3.metric("Total Penduduk", "850.057")
        col3.metric(" Laki Laki (Smokers)", "255.437")
        col4.metric("Jumlah Sektor", "13")
        col4.metric("Perkapita Rokok/Mgg (Rp)", "17.599")

    if Zona == 'Sumenep' :
        col1, col2, col3, col4 = st.columns(4)
    
        col1.metric("MP Jt Batang/Hari", "3,3")
        col1.metric("MP Ball Standard/Mgg", "8.367")
        col2.metric("OU", "13.635")
        col2.metric("Outlet Database", "2.493")
        col3.metric("Total Penduduk", "1.124.436")
        col3.metric(" Laki Laki (Smokers)", "328.483")
        col4.metric("Jumlah Sektor", "27")
        col4.metric("Perkapita Rokok/Mgg (Rp)", "16.719")

####################### Struktur Organisasi ####################

if option == "Struktur Organisasi":
    st.markdown("""---""")

    st.header("Struktur Organisasi")
image1 = Image.open('Team_Promosi.png')
image2 = Image.open('Team_Distribusi.png')
image3 = Image.open('Back_Office.png')

col1, col2, col3 = st.columns(3)
with col1 :
    if option == "Struktur Organisasi":
        st.subheader('Promotion Team')
        st.image(image1, caption='Promotion Team')
with col2 :
    if option == "Struktur Organisasi":
        st.subheader('Distribution Team')
        st.image(image2, caption='Distribution Team')
with col3 :
   if option == "Struktur Organisasi":
       st.subheader('Back Office')
       st.image(image3, caption='Back Office')

####################### MP By Zona - Kategori - Kelas ####################
st.markdown("""---""")
if option == "MP By Zona - Kategori - Kelas":
   
    st.markdown('### MP By ZONA')
    
    @st.cache(allow_output_mutation=True)
    def load_data():
        df = pd.read_excel(
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
            engine="openpyxl",
            sheet_name="MP_Zona",
            usecols="H:L",
            nrows=48,
            )
        return df

    df = load_data()

    fig = px.pie(df, values='MP_Ball_Standard', names='Zona',width= 500, height= 400, title="MP By Zona Per Jun 2022")
                        
    st.plotly_chart(fig) 

    st.markdown("""---""")

    st.markdown('### MP DSO By Kategori & Kelas')
    col1, col2 = st.columns(2)
    with col2 :
        if option == "MP By Zona - Kategori - Kelas":
                df = pd.read_excel(
                io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
                engine="openpyxl",
                sheet_name="Sektor",
                usecols="A:P",
                nrows=10000,
                )
                fig = px.pie(df, values='Jun22', names='Kategori-Kelas',width= 500, height= 400, title="MP By Kategori Rokok (Kelas) Per Jun 22")
                        
                st.plotly_chart(fig)

    with col1 :
        if option == "MP By Zona - Kategori - Kelas":
                df = pd.read_excel(
                io="/Users/ekasulawestara/Dashboard DSO Pamekasan/Market_DSO.xlsx",
                engine="openpyxl",
                sheet_name="MARKET",
                usecols="B:C",
                nrows=1000,
                )
                
                fig = px.pie(df, values='MP', names='M.Potensi_per_Kategori_Rokok',width= 500, height= 400, title="MP By Kategori Rokok 2022")
                
                st.plotly_chart(fig)

    
    @st.cache(allow_output_mutation=True)
    def load_data():
        df = pd.read_excel(
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
            engine="openpyxl",
            sheet_name="MP_Zona",
            usecols="H:L",
            nrows=48,
            )
        return df

    df = load_data()

    col1, col2 = st.columns(2)
    with col1 :
        if option == "MP By Zona - Kategori - Kelas":

            Zona = st.radio(
                "Select Zona",
                ('Bangkalan', 'Sampang', 'Pamekasan', 'Sumenep'), horizontal= True)
            if Zona == 'Bangkalan':
                
                table = pd.pivot_table( data=df, 
                                        index=['Kategori'], 
                                        columns=['Zona'], 
                                        values='MP_Ball_Standard',
                                        aggfunc='sum',
                                        fill_value=0
                                                )

                fig = px.pie(table, values=Zona, names= ['SKM', 'SKML','SKT','SPM'], width= 500, height= 400, title="MP Zona Bangkalan By Kategori Per Jun 2022")
                st.plotly_chart(fig)
    
                

            if Zona == 'Sampang':
                table = pd.pivot_table( data=df, 
                                        index=['Kategori'], 
                                        columns=['Zona'], 
                                        values='MP_Ball_Standard',
                                        aggfunc='sum',
                                        fill_value=0
                                                )
            
                fig = px.pie(table, values=Zona,names= ['SKM', 'SKML','SKT','SPM'],width= 500, height= 400, title="MP Zona Sampang By Kategori Per Jun 2022")
                st.plotly_chart(fig)

            if Zona == 'Pamekasan':
                table = pd.pivot_table( data=df, 
                                        index=['Kategori'], 
                                        columns=['Zona'], 
                                        values='MP_Ball_Standard',
                                        aggfunc='sum',
                                        fill_value=0
                                                )
                            
                fig = px.pie(table, values=Zona, names= ['SKM', 'SKML','SKT','SPM'], width= 500, height= 400, title="MP Zona Pamekasan By Kategori Per Jun 2022")
                st.plotly_chart(fig)

            if Zona == 'Sumenep' :
                table = pd.pivot_table( data=df, 
                                        index=['Kategori'], 
                                        columns=['Zona'], 
                                        values='MP_Ball_Standard',
                                        aggfunc='sum',
                                        fill_value=0
                                                )
                            
                fig = px.pie(table, values=Zona, names= ['SKM', 'SKML','SKT','SPM'], width= 500, height= 400, title="MP Zona Sumenep By Kategori Per Jun  2022")
                st.plotly_chart(fig)
                              
    with col2 :
        if option == "MP By Zona - Kategori - Kelas":
            st.markdown("""---""")
            
            if Zona == 'Bangkalan':
                table = pd.pivot_table( data=df, 
                                        index=['Kat-Kel'], 
                                        columns=['Zona'], 
                                        values='MP_Ball_Standard',
                                        aggfunc='sum',
                                        fill_value=0
                                                )
            
                fig = px.pie(table, values=Zona, names= ['SKM-Low','SKM-Med','SKM-Prem','SKML-Low','SKML-Med','SKML-Prem','SKT-Low','SKT-Med','SKT-Prem','SPM-Low','SPM-Med','SPM-Prem'], width= 500, height= 400, title="MP Zona Bangkalan By Kategori-Kelas Per Jun 2022")
                st.plotly_chart(fig)

            if Zona == 'Sampang':
                table = pd.pivot_table( data=df, 
                                        index=['Kat-Kel'], 
                                        columns=['Zona'], 
                                        values='MP_Ball_Standard',
                                        aggfunc='sum',
                                        fill_value=0
                                                )
            
                fig = px.pie(table, values=Zona, names= ['SKM-Low','SKM-Med','SKM-Prem','SKML-Low','SKML-Med','SKML-Prem','SKT-Low','SKT-Med','SKT-Prem','SPM-Low','SPM-Med','SPM-Prem'], width= 500, height= 400, title="MP Zona Sampang By Kategori-Kelas Per Jun 2022")
                st.plotly_chart(fig)
                
            if Zona == 'Pamekasan':
                table = pd.pivot_table( data=df, 
                                        index=['Kat-Kel'], 
                                        columns=['Zona'], 
                                        values='MP_Ball_Standard',
                                        aggfunc='sum',
                                        fill_value=0
                                                )
            
                fig = px.pie(table, values=Zona, names= ['SKM-Low','SKM-Med','SKM-Prem','SKML-Low','SKML-Med','SKML-Prem','SKT-Low','SKT-Med','SKT-Prem','SPM-Low','SPM-Med','SPM-Prem'], width= 500, height= 400, title="MP Zona Pamekasan By Kategori-Kelas Per Jun 2022")
                st.plotly_chart(fig)

            if Zona == 'Sumenep' :
                table = pd.pivot_table( data=df, 
                                        index=['Kat-Kel'], 
                                        columns=['Zona'], 
                                        values='MP_Ball_Standard',
                                        aggfunc='sum',
                                        fill_value=0
                                                )
            
                fig = px.pie(table, values=Zona, names= ['SKM-Low','SKM-Med','SKM-Prem','SKML-Low','SKML-Med','SKML-Prem','SKT-Low','SKT-Med','SKT-Prem','SPM-Low','SPM-Med','SPM-Prem'], width= 500, height= 400, title="MP Zona Sumenep By Kategori-Kelas Per Jun 2022")
                st.plotly_chart(fig)

####################### MS By Group Pabrikan ####################

if option == "MS By Group Pabrikan":
    col1, col2 = st.columns(2)
    with col1 :
        st.markdown('### MS By Group - Nielsen (Per Mei 2022)')
        fig = plt.figure(figsize=(9, 5.0625))
        ax1 = fig.add_subplot(121)
        ax2 = fig.add_subplot(122)
        fig.subplots_adjust(wspace=0)
        
        # large pie chart parameters
        
        df = pd.read_excel(
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/Market_DSO.xlsx",
            engine="openpyxl",
            sheet_name="Nielsen",
            usecols="A:P",
            nrows=1000,
            )
        pivot1 = pd.pivot_table(data=df, index=['Group'], values = 'May-22', aggfunc='sum')
        xxx = [pivot1.iloc[0]['May-22'], pivot1.iloc[1]['May-22'], pivot1.iloc[2]['May-22'], pivot1.iloc[3]['May-22'], pivot1.iloc[4]['May-22'],pivot1.iloc[5]['May-22']]
        pivot2 = pd.pivot_table(data=df, index=['Group', 'Kategori'], values = 'May-22', aggfunc='sum')
        yyy = [pivot2.iloc[4]['May-22'], pivot2.iloc[5]['May-22'], pivot2.iloc[6]['May-22']]

        ratios = xxx
        labels = ['BAT', 'Djarum', 'GG', 'Nojorono', 'Others', 'PMI']
        explode = [0, 0.1, 0, 0, 0, 0]
        # rotate so that first wedge is split by the x-axis
        angle = -60 * ratios[0]
        ax1.pie(ratios, autopct='%1.1f%%', startangle=angle,
                labels=labels, explode=explode)
        # small pie chart parameters
        ratios = yyy
        labels = ['SKM', 'SKML', 'SKT']
        width = .2
        ax2.pie(ratios, autopct='%1.1f%%', startangle=angle,
                labels=labels, radius=0.5, textprops={'size': 'smaller'})
        ax1.set_title('MS by Group')
        ax2.set_title('Kontribusi by Kategori')
        # use ConnectionPatch to draw lines between the two plots
        # get the wedge data
        theta1, theta2 = ax1.patches[1].theta1, ax1.patches[1].theta2
        center, r = ax1.patches[1].center, ax1.patches[1].r
        # draw top connecting line
        x = r * np.cos(np.pi / 180 * theta2) + center[0]
        y = np.sin(np.pi / 180 * theta2) + center[1]
        con = ConnectionPatch(xyA=(- width / 2, .5), xyB=(x, y),
                            coordsA="data", coordsB="data", axesA=ax2, axesB=ax1)
        con.set_color([0, 0, 0])
        con.set_linewidth(0.5)
        ax2.add_artist(con)
        # draw bottom connecting line
        x = r * np.cos(np.pi / 180 * theta1) + center[0]
        y = np.sin(np.pi / 180 * theta1) + center[1]
        con = ConnectionPatch(xyA=(- width / 2, -.5), xyB=(x, y), coordsA="data",
                            coordsB="data", axesA=ax2, axesB=ax1)
        con.set_color([0, 0, 0])
        ax2.add_artist(con)
        con.set_linewidth(0.5)
        st.write(fig)


    with col2 :    

        st.markdown('### MS By Group - Internal (Per Juni 2022)')
        fig = plt.figure(figsize=(9, 5.0625))
        ax1 = fig.add_subplot(121)
        ax2 = fig.add_subplot(122)
        fig.subplots_adjust(wspace=0)
        
        # large pie chart parameters
        
        @st.cache(allow_output_mutation=True)
        def load_data():
            df = pd.read_excel(
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
            engine="openpyxl",
            sheet_name="DSO",
            usecols="A:R",
            nrows=200,
            )
            return df

        df = load_data()
        
        pivot1 = pd.pivot_table(data=df, index=['Group'], values = 'Jun22', aggfunc='sum')
        xxx1 = [pivot1.iloc[0]['Jun22'], pivot1.iloc[1]['Jun22'], pivot1.iloc[2]['Jun22'], pivot1.iloc[3]['Jun22'], pivot1.iloc[4]['Jun22'], pivot1.iloc[5]['Jun22']]
        pivot2 = pd.pivot_table(data=df, index=['Kategori', 'Group'], values = 'Jun22', aggfunc='sum')
        yyy1 = [pivot2.iloc[1]['Jun22'], pivot2.iloc[6]['Jun22'], pivot2.iloc[11]['Jun22']]
        ratios = xxx1
        labels = ['NTI', 'Djarum', 'GG', 'BAT', 'Others', 'PMI']
        explode = [0, 0.1, 0, 0, 0, 0]
        # rotate so that first wedge is split by the x-axis
        angle = -33 * ratios[0]
        ax1.pie(ratios, autopct='%1.1f%%', startangle=angle,
                labels=labels, explode=explode)
        # small pie chart parameters
        ratios = yyy1
        labels = ['SKM', 'SKML', 'SKT']
        width = .2
        ax2.pie(ratios, autopct='%1.1f%%', startangle=angle,
                labels=labels, radius=0.5, textprops={'size': 'smaller'})
        ax1.set_title('MS by Group')
        ax2.set_title('Kontribusi by Kategori')
        # use ConnectionPatch to draw lines between the two plots
        # get the wedge data
        theta1, theta2 = ax1.patches[1].theta1, ax1.patches[1].theta2
        center, r = ax1.patches[1].center, ax1.patches[1].r
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

#################################### MS Zona By Group ############################################

if option == "MS By Group Pabrikan":
        Zona = st.radio(
            "Select Zona",
            ('All Zona', 'Bangkalan', 'Sampang', 'Pamekasan', 'Sumenep'), horizontal= True)

        df_dso = pd.read_excel(
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
            engine="openpyxl",
            sheet_name="DSO",
            usecols="A:T",
            nrows=200,
            )
    
        if Zona == 'All Zona':
                
                table_dso = pd.pivot_table( data=df_dso, 
                                        index=['Group','Kategori'], 
                                        columns=['DSO'], 
                                        values='Jun22',
                                        aggfunc='sum',
                                        fill_value=0
                                                )
                fig = px.pie(table_dso, values='Pamekasan', names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by All Zona Per Jun 2022")
                st.plotly_chart(fig)


        df = pd.read_excel(
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
            engine="openpyxl",
            sheet_name="Zona",
            usecols="A:N",
            nrows=500,
            )

        if Zona == 'Bangkalan':

                    table = pd.pivot_table( data=df, 
                                            index=['Group','Kategori'], 
                                            columns=['Zona'], 
                                            values='Jun22',
                                            aggfunc='sum',
                                            fill_value=0
                                                    )

                    fig = px.pie(table, values=Zona, names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by Zona Bangkalan Per Jun 2022")
                    st.plotly_chart(fig)

        if Zona == 'Sampang':
                    table = pd.pivot_table( data=df, 
                                            index=['Group','Kategori'], 
                                            columns=['Zona'], 
                                            values='Jun22',
                                            aggfunc='sum',
                                            fill_value=0
                                                    )

                    fig = px.pie(table, values=Zona, names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by Zona Sampang Per Jun 2022")
                    st.plotly_chart(fig)

        if Zona == 'Pamekasan':
                    table = pd.pivot_table( data=df, 
                                            index=['Group','Kategori'], 
                                            columns=['Zona'], 
                                            values='Jun22',
                                            aggfunc='sum',
                                            fill_value=0
                                                    )

                    fig = px.pie(table, values=Zona, names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by Zona Pamekasan Per Jun 2022")
                    st.plotly_chart(fig)

        if Zona == 'Sumenep' :
                    table = pd.pivot_table( data=df, 
                                            index=['Group','Kategori'], 
                                            columns=['Zona'], 
                                            values='Jun22',
                                            aggfunc='sum',
                                            fill_value=0
                                                    )

                    fig = px.pie(table, values=Zona, names= ['BAT Group','BAT Group','BAT Group', 'Djarum Group', 'Djarum Group', 'Djarum Group','GG Group','GG Group','GG Group','GG Group','NTI Group','NTI Group','Others Group','Others Group','Others Group','Others Group','PMI Group','PMI Group','PMI Group','PMI Group'], width= 500, height= 400, title="MS Group Pabrikan by Zona Sumenep Per Jun 2022")
                    st.plotly_chart(fig)
        
        ######### MS Zona By Group Kategori ########

        if st.button ("Show Data"):  

            if Zona =='All Zona':
                col1, col2 = st.columns(2)
                with col1 :        
                        table_dso = pd.pivot_table( data=df_dso, 
                                                index=['Group-Kat'], 
                                                columns=['DSO'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table_dso, values='Pamekasan', names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by All Zona Per Jun 2022")
                        st.plotly_chart(fig)

                with col2 :
                        table_dso = pd.pivot_table( data=df_dso, 
                                                index=['Group-Kat-Kel'], 
                                                columns=['DSO'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table_dso, values='Pamekasan', names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by All Zona Per Jun 2022")
                        st.plotly_chart(fig)

        
            if Zona == 'Bangkalan':

                col1, col2 = st.columns(2)
                with col1 :        
                        table = pd.pivot_table( data=df, 
                                                index=['Group-Kat'], 
                                                columns=['Zona'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by Zona Bangkalan Per Jun 2022")
                        st.plotly_chart(fig)

                with col2 :
                        table = pd.pivot_table( data=df, 
                                                index=['Group-Kat-Kel'], 
                                                columns=['Zona'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by Zona Bangkalan Per Jun 2022")
                        st.plotly_chart(fig)


            if Zona == 'Sampang':

                col1, col2 = st.columns(2)
                with col1 :  
                        table = pd.pivot_table( data=df, 
                                                index=['Group-Kat'], 
                                                columns=['Zona'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by Zona Sampang Per Jun 2022")
                        st.plotly_chart(fig)

                with col2 :
                        table = pd.pivot_table( data=df, 
                                                index=['Group-Kat-Kel'], 
                                                columns=['Zona'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by Zona Sampang Per Jun 2022")
                        st.plotly_chart(fig)

            if Zona == 'Pamekasan':

                col1, col2 = st.columns(2)
                with col1 :  
                        table = pd.pivot_table( data=df, 
                                                index=['Group-Kat'], 
                                                columns=['Zona'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by Zona Pamekasan Per Jun 2022")
                        st.plotly_chart(fig)

                with col2 :
                        table = pd.pivot_table( data=df, 
                                                index=['Group-Kat-Kel'], 
                                                columns=['Zona'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by Zona Pamekasan Per Jun 2022")
                        st.plotly_chart(fig)

            if Zona == 'Sumenep' :

                col1, col2 = st.columns(2)
                with col1 :  
                        table = pd.pivot_table( data=df, 
                                                index=['Group-Kat'], 
                                                columns=['Zona'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM','BAT Group-SKML','BAT Group-SPM', 'Djarum Group-SKM', 'Djarum Group-SKML', 'Djarum Group-SKT','GG Group-SKM','GG Group-SKML','GG Group-SKT','GG Group-SPM','NTI Group-SKML','NTI Group-SKT','Others Group-SKM','Others Group-SKML','Others Group-SKT','Others Group-SPM','PMI Group-SKM','PMI Group-SKML','PMI Group-SKT','PMI Group-SPM'], width= 500, height= 400, title="MS Group Pabrikan Kategori by Zona Sumenep Per Jun 2022")
                        st.plotly_chart(fig)

                with col2 :
                        table = pd.pivot_table( data=df, 
                                                index=['Group-Kat-Kel'], 
                                                columns=['Zona'], 
                                                values='Jun22',
                                                aggfunc='sum',
                                                fill_value=0
                                                        )

                        fig = px.pie(table, values=Zona, names= ['BAT Group-SKM-Prem','BAT Group-SKML-Prem', 'BAT Group-SPM-Low', 'BAT Group-SPM-Prem', 'Djarum Group-SKM-Low', 'Djarum Group-SKM-Med', 'Djarum Group-SKM-Prem', 'Djarum Group-SKML-Low', 'Djarum Group-SKML-Prem', 'Djarum Group-SKT-Low','Djarum Group-SKT-Med', 'GG Group-SKM-Low', 'GG Group-SKM-Med', 'GG Group-SKM-Prem', 'GG Group-SKML-Prem', 'GG Group-SKT-Low', 'GG Group-SKT-Med', 'GG Group-SPM-Low', 'NTI Group-SKML-Prem', 'NTI Group-SKT-Low', 'Others Group-SKM-Low', 'Others Group-SKM-Prem', 'Others Group-SKML-Low', 'Others Group-SKML-Med', 'Others Group-SKML-Prem', 'Others Group-SKT-Low', 'Others Group-SKT-Med', 'Others Group-SKT-Prem', 'Others Group-SPM-Low', 'Others Group-SPM-Prem', 'PMI Group-SKM-Med', 'PMI Group-SKM-Prem', 'PMI Group-SKML-Prem', 'PMI Group-SKT-Low','PMI Group-SKT-Med', 'PMI Group-SKT-Prem', 'PMI Group-SPM-Prem'], width= 500, height= 400, title="MS Group Pabrikan Kategori-Kelas by Zona Sumenep Per Jun 2022")
                        st.plotly_chart(fig)

            st.button ("Close Data")


############ Top 10 Brand Nielsen #############

if option == "Top 10 Brand Nielsen" :

    st.markdown('### Top 10 Brand MS Nielsen - Mei 2022 (YTD)')

    df = pd.read_excel(
    io="/Users/ekasulawestara/Dashboard DSO Pamekasan/Market_DSO.xlsx",
    engine="openpyxl",
    sheet_name="Nielsen",
    usecols="A:Q",
    nrows=10
    )
    new_df = df.drop(['Group','Kategori'], axis=1)

    first =((new_df).iloc[0,0])
    first_value = ('{:.2f}%'.format((new_df).iloc[0,13]))
    first_delta = ('{:.2f}%'.format((new_df).iloc[0,14])) 

    second =((new_df).iloc[1,0])
    second_value = ('{:.2f}%'.format((new_df).iloc[1,13]))
    second_delta = ('{:.2f}%'.format((new_df).iloc[1,14]))  

    thirth =((new_df).iloc[2,0])
    thirth_value = ('{:.2f}%'.format((new_df).iloc[2,13]))
    thirth_delta = ('{:.2f}%'.format((new_df).iloc[2,14]))

    fourth =((new_df).iloc[3,0])
    fourth_value = ('{:.2f}%'.format((new_df).iloc[3,13]))
    fourth_delta = ('{:.2f}%'.format((new_df).iloc[3,14]))

    fifth =((new_df).iloc[4,0])
    fifth_value = ('{:.2f}%'.format((new_df).iloc[4,13]))
    fifth_delta = ('{:.2f}%'.format((new_df).iloc[4,14]))

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric(first,first_value, first_delta)
    col2.metric(second,second_value, second_delta)
    col3.metric(thirth, thirth_value, thirth_delta)
    col4.metric(fourth, fourth_value, fourth_delta)
    col5.metric(fifth, fifth_value, fifth_delta)

    sixth =((new_df).iloc[5,0])
    sixth_value = ('{:.2f}%'.format((new_df).iloc[5,13]))
    sixth_delta = ('{:.2f}%'.format((new_df).iloc[5,14])) 
    seventh =((new_df).iloc[6,0])
    seventh_value = ('{:.2f}%'.format((new_df).iloc[6,13]))
    seventh_delta = ('{:.2f}%'.format((new_df).iloc[6,14]))  
    eighth =((new_df).iloc[7,0])
    eighth_value = ('{:.2f}%'.format((new_df).iloc[7,13]))
    eighth_delta = ('{:.2f}%'.format((new_df).iloc[7,14]))
    ninth =((new_df).iloc[8,0])
    ninth_value = ('{:.2f}%'.format((new_df).iloc[8,13]))
    ninth_delta = ('{:.2f}%'.format((new_df).iloc[8,14]))
    tenth =((new_df).iloc[9,0])
    tenth_value = ('{:.2f}%'.format((new_df).iloc[9,13]))
    tenth_delta = ('{:.2f}%'.format((new_df).iloc[9,14]))

    col6, col7, col8, col9, col10 = st.columns(5)
    col6.metric(sixth,sixth_value, sixth_delta)
    col7.metric(seventh,seventh_value, seventh_delta)
    col8.metric(eighth, eighth_value, eighth_delta)
    col9.metric(ninth, ninth_value, ninth_delta)
    col10.metric(tenth, tenth_value, tenth_delta)

    st.markdown("""---""")

    with st.expander("Show Data (Magnificent 7 Brand Prioritas - Djarum)"):

        names = ['LA Bold', 'Geo Mild', 'Chief Filter','L.A. Lights Regular','D. 7 6', 'D. Super','Chief Kretek']

        df2 = pd.read_excel(
        io="/Users/ekasulawestara/Dashboard DSO Pamekasan/Market_DSO.xlsx",
        engine="openpyxl",
        sheet_name="Nielsen",
        usecols="A:Q",
        nrows=100
        )
        df_djarum_group = df2[df2.Brand.isin(names)]
        df_top_10_internal = df_djarum_group.drop(['Group','Kategori','Delta',], axis=1)
        st.write(df_top_10_internal)

############ Top 10 Brand Internal #############

if option == "Top 10 Brand Internal" :

        st.markdown('### Top 10 Brand MS Internal - Juni 2022 (YTD)')
        
        names = ['LA Bold', 'Geo Mild', 'Chief Filter','L.A. Lights Regular','D. 7 6', 'D. Super','Chief Kretek']

        df = pd.read_excel(
        io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
        engine="openpyxl",
        sheet_name="DSO",
        usecols="A:R",
        nrows=10
        )
        new_df = df.drop(['Group','Kategori', 'Kelas','DSO','Group-Kat','Group-Kat-Kel','Agt22','Okt22', 'Des22'], axis=1)

        first =((new_df).iloc[0,0])
        first_value = ('{:.2f}%'.format((new_df).iloc[0,7]))
        first_delta = ('{:.2f}%'.format((new_df).iloc[0,8])) 

        second =((new_df).iloc[1,0])
        second_value = ('{:.2f}%'.format((new_df).iloc[1,7]))
        second_delta = ('{:.2f}%'.format((new_df).iloc[1,8]))  

        thirth =((new_df).iloc[2,0])
        thirth_value = ('{:.2f}%'.format((new_df).iloc[2,7]))
        thirth_delta = ('{:.2f}%'.format((new_df).iloc[2,8]))

        fourth =((new_df).iloc[3,0])
        fourth_value = ('{:.2f}%'.format((new_df).iloc[3,7]))
        fourth_delta = ('{:.2f}%'.format((new_df).iloc[3,8]))

        fifth =((new_df).iloc[4,0])
        fifth_value = ('{:.2f}%'.format((new_df).iloc[4,7]))
        fifth_delta = ('{:.2f}%'.format((new_df).iloc[4,8]))

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric(first,first_value, first_delta)
        col2.metric(second,second_value, second_delta)
        col3.metric(thirth, thirth_value, thirth_delta)
        col4.metric(fourth, fourth_value, fourth_delta)
        col5.metric(fifth, fifth_value, fifth_delta)

        sixth =((new_df).iloc[5,0])
        sixth_value = ('{:.2f}%'.format((new_df).iloc[5,7]))
        sixth_delta = ('{:.2f}%'.format((new_df).iloc[5,8])) 
        seventh =((new_df).iloc[6,0])
        seventh_value = ('{:.2f}%'.format((new_df).iloc[6,7]))
        seventh_delta = ('{:.2f}%'.format((new_df).iloc[6,8]))  
        eighth =((new_df).iloc[7,0])
        eighth_value = ('{:.2f}%'.format((new_df).iloc[7,7]))
        eighth_delta = ('{:.2f}%'.format((new_df).iloc[7,8]))
        ninth =((new_df).iloc[8,0])
        ninth_value = ('{:.2f}%'.format((new_df).iloc[8,7]))
        ninth_delta = ('{:.2f}%'.format((new_df).iloc[8,8]))
        tenth =((new_df).iloc[9,0])
        tenth_value = ('{:.2f}%'.format((new_df).iloc[9,7]))
        tenth_delta = ('{:.2f}%'.format((new_df).iloc[9,8]))

        col6, col7, col8, col9, col10 = st.columns(5)
        col6.metric(sixth,sixth_value, sixth_delta)
        col7.metric(seventh,seventh_value, seventh_delta)
        col8.metric(eighth, eighth_value, eighth_delta)
        col9.metric(ninth, ninth_value, ninth_delta)
        col10.metric(tenth, tenth_value, tenth_delta)

        st.markdown("""---""")
        with st.expander("Show Data (Magnificent 7 Brand Prioritas - Djarum)"):

            df2 = pd.read_excel(
            io="Data/MIT.xlsx",
            engine="openpyxl",
            sheet_name="DSO",
            usecols="A:T",
            nrows=45
            )

            df_djarum_group = df2[df2.GabBrand.isin(names)]
            df_top_10_internal = df_djarum_group.drop(['DSO','Group','Kategori','Group-Kat', 'Group-Kat-Kel', 'Kelas','Delta', 'Agt22', 'Okt22', 'Des22'], axis=1)
            st.write(df_top_10_internal)
            

############ DSO Sales Performance #############

if option == "DSO Sales Performance":
    st.subheader(" SALES PERFORMACE")

    @st.cache(allow_output_mutation=True)
    def load_data():
        df10 = pd.read_excel(
            io="Data/MIT.xlsx",
            engine="openpyxl",
            sheet_name="Omset_2022",
            usecols="G:I",
            nrows=250000
            )
        return (df10)
    df10=load_data()

    df11 = df10.pivot_table('Omzet',index='RokokName_S',columns='TimePeriode',aggfunc={'Omzet':'sum'})
    df12 = df11.T.assign(TotalOmset = lambda x: x.sum(axis=1))
    df13 = df12[['TotalOmset']]
    
    fig = px.line(df13, title='Omset DSO Pamekasan Sem 1 2022')
    st.plotly_chart(fig, use_container_width=True)

        ### TOP 10 Brand##########

    df20 = df10.pivot_table('Omzet',index='RokokName_S',columns='TimePeriode',aggfunc={'Omzet':'sum'}).T
    df21 = df20.T.assign(Rata2 = lambda x: x.mean(axis=1))
    df22 = df21.sort_values('Rata2',ascending = False).groupby('RokokName_S').head(2)
    df23 = df22.head(10).T
    df24 = df23.head(23)

    fig = px.line(df24, title='Top 10 Brand')
    st.plotly_chart(fig, use_container_width=True)

    @st.cache(allow_output_mutation=True)
    def load_data():
        df40 = pd.read_excel(
            io="Data/MIT.xlsx",
            engine="openpyxl",
            sheet_name="Omset_2022",
            usecols="E:I",
            nrows=250000
            )
        return (df40)
    df40=load_data()

    df41 = df40.drop(['RokokProductGroup','RokokName_S'], axis=1)
    df42 = df41.pivot_table('Omzet',index='GeographyZona',columns='TimePeriode',aggfunc={'Omzet':'sum'}).T    
    
    fig = px.line(df42, title='Kontribusi Omset By Zona')
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("""---""")

if option == "DSO Sales Performance":

    @st.cache(allow_output_mutation=True)
    def load_data():
        df = pd.read_excel(
            io="Data/Market_DSO.xlsx",
            engine="openpyxl",
            sheet_name="MARKET",
            usecols="B:E",
            nrows=1000,
        )
        return (df)
    df=load_data()

    @st.cache(allow_output_mutation=True)
    def load_data():
        df = pd.read_excel(
            io="Data/MIT.xlsx",
            engine="openpyxl",
            sheet_name="Sektor",
            usecols="A:P",
            nrows=10000,
        )
        return (df)
    df=load_data()

    @st.cache(allow_output_mutation=True)
    def load_data():    
        df = pd.read_excel(
            io="Data/MIT.xlsx",
            engine="openpyxl",
            sheet_name="Omset_2022",
            usecols="A:J",
            nrows=250000,
            )
        return (df)
    df=load_data()
    
    if option == "DSO Sales Performance":
        col1, col2 = st.columns(2)
    with col1 :

            fig = px.pie(df, values='OmzetJtBtg', names='GeographyZona',width= 500, height= 400, title="Kontribusi Omset By Geography Zona Sem 1 2022")
    
            st.plotly_chart(fig)

    with col2 :
            fig = px.pie(df, values='OmzetJtBtg', names='RokokProductGroup',width= 500, height= 400, title="Kontribusi Omset By Group Kategori Rokok Sem 1 2022")
    
            st.plotly_chart(fig)

    st.markdown("""---""")

    if option == "DSO Sales Performance":
        st.markdown('### Kontribusi Omset By Customer Category Sem 1 2022')
        col1, col2 = st.columns(2)
    with col1 :

        fig = px.pie(df, values='OmzetJtBtg', names='CustomerGroupCategory',width= 500, height= 400, title="Kontribusi Omset By Customer Group Category Sem 1 2022")
        
        st.plotly_chart(fig)
    with col2 :

        fig = px.pie(df, values='OmzetJtBtg', names='CustomerCategory',width= 500, height= 400, title="Kontribusi Omset By CustomerCategory Sem 1 2022")
        
        st.plotly_chart(fig)  

if option == "DSO Sales Performance":
    st.markdown("""---""")

    st.markdown('### Kontribusi Omset Shop By Zona Sem 1 2022')
    fig = plt.figure(figsize=(12, 5.0625))
    ax1 = fig.add_subplot(121)
    ax2 = fig.add_subplot(122)
    fig.subplots_adjust(wspace=0)
    
    # large pie chart parameters

    pivot1 = pd.pivot_table(data=df, index=['CustomerGroupCategory'], values = 'OmzetJtBtg', aggfunc='sum')
    xxx1 = [pivot1.iloc[1]['OmzetJtBtg'], pivot1.iloc[0]['OmzetJtBtg'], pivot1.iloc[2]['OmzetJtBtg'], pivot1.iloc[3]['OmzetJtBtg']]
    pivot2 = pd.pivot_table(data=df, index=['GeographyZona', 'CustomerGroupCategory'], values = 'OmzetJtBtg', aggfunc='sum')
    yyy1 = [pivot2.iloc[1]['OmzetJtBtg'], pivot2.iloc[5]['OmzetJtBtg'], pivot2.iloc[9]['OmzetJtBtg'], pivot2.iloc[13]['OmzetJtBtg']]

    ratios = xxx1
    labels = ['Shop', 'Modern Retail', 'Small Retail', 'Special Retail']
    explode = [0.1, 0, 0, 0]
    # rotate so that first wedge is split by the x-axis
    angle = 95 * ratios[0]
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
    ax2.set_title('Kontribusi Zona')
    # use ConnectionPatch to draw lines between the two plots
    # get the wedge data
    theta1, theta2 = ax1.patches[0].theta1, ax1.patches[0].theta2
    center, r = ax1.patches[0].center, ax1.patches[0].r
    # draw top connecting line
    x = r * np.cos(np.pi / -90 * theta2) + center[0]
    y = np.sin(np.pi / -90 * theta2) + center[1]
    con = ConnectionPatch(xyA=(- width / 2, .5), xyB=(x, y),
                        coordsA="data", coordsB="data", axesA=ax2, axesB=ax1)
    con.set_color([0, 0, 0])
    con.set_linewidth(.5)
    ax2.add_artist(con)
    # draw bottom connecting line
    x = r * np.cos(np.pi / -100 * theta1) + center[0]
    y = np.sin(np.pi / -100 * theta1) + center[1]
    con = ConnectionPatch(xyA=(- width / 2, -.5), xyB=(x, y), coordsA="data",
                        coordsB="data", axesA=ax2, axesB=ax1)
    con.set_color([0, 0, 0])
    ax2.add_artist(con)
    con.set_linewidth(.5)
    st.write(fig)  

############ DSO Canvas Performance #############
  
if option == "DSO Canvas Performance":

    st.markdown('### Kontribusi Omset Kanvas by Zona Sem 1 2022')
    fig = plt.figure(figsize=(12, 5.0625))
    ax1 = fig.add_subplot(121)
    ax2 = fig.add_subplot(122)
    fig.subplots_adjust(wspace=0)
    
    @st.cache(allow_output_mutation=True)
    def load_data():    
        df = pd.read_excel(
            io="Data/MIT.xlsx",
            engine="openpyxl",
            sheet_name="Omset_2022",
            usecols="A:J",
            nrows=250000,
            )
        return (df)
    df=load_data()

    # large pie chart parameters 

    pivot1 = pd.pivot_table(data=df, index=['CustomerGroupCategory'], values = 'OmzetJtBtg', aggfunc='sum')
    xxx1 = [pivot1.iloc[1]['OmzetJtBtg'], pivot1.iloc[0]['OmzetJtBtg'], pivot1.iloc[2]['OmzetJtBtg'], pivot1.iloc[3]['OmzetJtBtg']]
    pivot2 = pd.pivot_table(data=df, index=['GeographyZona', 'CustomerGroupCategory'], values = 'OmzetJtBtg', aggfunc='sum')
    yyy1 = [pivot2.iloc[2]['OmzetJtBtg'], pivot2.iloc[6]['OmzetJtBtg'], pivot2.iloc[10]['OmzetJtBtg'], pivot2.iloc[14]['OmzetJtBtg']]

    ratios = xxx1
    labels = ['Shop', 'Modern Retail', 'Small Retail', 'Special Retail']
    explode = [0, 0, 0.1, 0]
    # rotate so that first wedge is split by the x-axis
    angle = -10.55 * ratios[0]
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

    st.markdown('### TDN NIELSEN')

    names = ['LA Bold', 'Geo Mild', 'Chief Filter','L.A. Lights Regular','D. 7 6', 'D. Super','Chief Kretek','D. 7 6 Madu Hitam', 'Gaze Kretek','Ferro Filter']

    df = pd.read_excel(
    io="Data/Market_DSO.xlsx",
    engine="openpyxl",
    sheet_name="TDN_Nielsen",
    usecols="A:Q",
    nrows=300
    )

    df1 = df.iloc[:,:1].T ###Header Brand
    df2 = df.drop(['Brand','Group', 'Kategori','Kelas'], axis=1).T ###Body Data
    df3 = df2.applymap(str)

    df4 = pd.concat([df1, df3], ignore_index = True, axis = True)
    df5 = df1.append(df3)
    df6 = df5.iloc[:, 0:10]
    df6.columns = df6.iloc[0]
    df6 = df6[1:]
    df7 = df6.applymap(int)
    ###st.write(df7)    
    fig = px.line(df7, title='TDN Nielsen Top 10 Brand')

    st.plotly_chart(fig, use_container_width=True)
    ###st.write(df7)

    df8 = df[df.Brand.isin(names)]
    df9 = df8.drop(['Group','Kategori','Kelas'], axis=1)
    df10= df9


    df11 = df10.iloc[:,:1].T ###Header Brand
    df12 = df9.drop(['Brand'], axis=1).T ###Body Data
    df13 = df12.applymap(str)

    df14 = pd.concat([df11, df13], ignore_index = True, axis = True)
    df15 = df11.append(df13)
    df16 = df15.iloc[:, 0:10]
    df16.columns = df16.iloc[0]
    df16 = df16[1:]
    df17 = df16.applymap(int)

    fig = px.line(df17, title='TDN Nielsen Top 10 Brand Djarum')

    st.plotly_chart(fig, use_container_width=True)
