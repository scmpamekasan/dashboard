from functools import cache
import pandas as pd  # pip install pandas openpyxl
import plotly as pl
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
import time
import requests
import json



st.set_page_config(layout="wide")

####################### HEADER ####################
st.header('DASHBOARD DSO PAMEKASAN')

#################### Menu SideBar ##################


option = st.sidebar.selectbox("Menu",["Home Page", "Struktur Organisasi", "MS & MP By Category", "MS By Group Pabrikan", "Top 10 Brand Nielsen", "Top 10 Brand Internal","DSO Sales Performance"])

####################### MAP DSO ####################
if option == "Home Page":
    st.markdown("""---""")

    height = 200
    width = 1300
    
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
    ###############################


    ####################### MAIN PAGE ####################


    col1, col2, col3, col4 = st.columns(4)
    col1.metric("MP Jt Batang", "12,4")
    col2.metric("OU", "48.560")
    col3.metric("Total Penduduk", "4.004.563")
    col4.metric("Jumlah Zona", "4")


    col1, col2, col3, col4 = st.columns(4)
    col1.metric("MP Ball Standard", "30.880")
    col2.metric("Outlet Database", "9.XXX")
    col3.metric(" Laki Laki (Smokers)", "1.212.375")
    col4.metric("Jumlah Sektor", "72")

    st.markdown("""---""")




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

####################### MS & MP By Category ####################

if option == "MS & MP By Category":
    
    st.markdown("""---""")

    st.markdown('### MS & MP By Category')
    col1, col2 = st.columns(2)
    with col2 :
        if option == "MS & MP By Category":
                df = pd.read_excel(
                io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
                engine="openpyxl",
                sheet_name="Sektor",
                usecols="A:P",
                nrows=10000,
                )
                fig = px.pie(df, values='Jun22', names='Kategori-Kelas',width= 500, height= 400, title="MS By Kategori Rokok (Kelas) Per Jun 22")
                        
                st.plotly_chart(fig)

    with col1 :
        if option == "MS & MP By Category":
                df = pd.read_excel(
                io="/Users/ekasulawestara/Dashboard DSO Pamekasan/Market_DSO.xlsx",
                engine="openpyxl",
                sheet_name="MARKET",
                usecols="B:C",
                nrows=1000,
                )
                
                fig = px.pie(df, values='MP', names='M.Potensi_per_Kategori_Rokok',width= 500, height= 400, title="MP By Kategori Rokok 2022")
                
                st.plotly_chart(fig)


####################### MS By Group Pabrikan ####################

if option == "MS By Group Pabrikan":
    st.markdown("""---""")

    st.markdown('### MS By Group - Nielsen (Per April 2022)')
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
    pivot1 = pd.pivot_table(data=df, index=['Group'], values = 'Apr-22', aggfunc='sum')
    xxx = [pivot1.iloc[0]['Apr-22'], pivot1.iloc[1]['Apr-22'], pivot1.iloc[2]['Apr-22'], pivot1.iloc[3]['Apr-22'], pivot1.iloc[4]['Apr-22']]
    pivot2 = pd.pivot_table(data=df, index=['Kategori'], values = 'Apr-22', aggfunc='sum')
    yyy = [pivot2.iloc[1]['Apr-22'], pivot2.iloc[2]['Apr-22'], pivot2.iloc[3]['Apr-22']]
    ratios = xxx
    labels = ['Djarum', 'GG', 'NTI', 'Others', 'PMI']
    explode = [0.1, 0, 0, 0, 0]
    # rotate so that first wedge is split by the x-axis
    angle = 65 * ratios[0]
    ax1.pie(ratios, autopct='%1.1f%%', startangle=angle,
            labels=labels, explode=explode)
    # small pie chart parameters
    ratios = yyy
    labels = ['SKM', 'SKML', 'SKT']
    width = .2
    ax2.pie(ratios, autopct='%1.1f%%', startangle=angle,
            labels=labels, radius=0.5, textprops={'size': 'smaller'})
    ax1.set_title('MS by Group')
    ax2.set_title('Djarum by Kategori')
    # use ConnectionPatch to draw lines between the two plots
    # get the wedge data
    theta1, theta2 = ax1.patches[0].theta1, ax1.patches[0].theta2
    center, r = ax1.patches[0].center, ax1.patches[0].r
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

if option == "MS By Group Pabrikan":
    st.markdown("""---""")

    st.markdown('### MS By Group - Internal (Per Juni 2022)')
    fig = plt.figure(figsize=(9, 5.0625))
    ax1 = fig.add_subplot(121)
    ax2 = fig.add_subplot(122)
    fig.subplots_adjust(wspace=0)
    
    # large pie chart parameters
    
    
    df = pd.read_excel(
        io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
        engine="openpyxl",
        sheet_name="Sektor",
        usecols="A:P",
        nrows=10000,
        )

   

    pivot1 = pd.pivot_table(data=df, index=['Group'], values = 'Jun22', aggfunc='sum')
    xxx1 = [pivot1.iloc[0]['Jun22'], pivot1.iloc[1]['Jun22'], pivot1.iloc[2]['Jun22'], pivot1.iloc[3]['Jun22'], pivot1.iloc[4]['Jun22'], pivot1.iloc[5]['Jun22']]
    pivot2 = pd.pivot_table(data=df, index=['Kategori', 'Group'], values = 'Jun22', aggfunc='sum')
    yyy1 = [pivot2.iloc[1]['Jun22'], pivot2.iloc[6]['Jun22'], pivot2.iloc[11]['Jun22']]
    ratios = xxx1
    labels = ['NTI', 'Djarum', 'GG', 'BAT', 'Others', 'PMI']
    explode = [0, 0.1, 0, 0, 0, 0]
    # rotate so that first wedge is split by the x-axis
    angle = -320.5 * ratios[0]
    ax1.pie(ratios, autopct='%1.1f%%', startangle=angle,
            labels=labels, explode=explode)
    # small pie chart parameters
    ratios = yyy1
    labels = ['SKM', 'SKML', 'SKT']
    width = .2
    ax2.pie(ratios, autopct='%1.1f%%', startangle=angle,
            labels=labels, radius=0.5, textprops={'size': 'smaller'})
    ax1.set_title('MS by Group')
    ax2.set_title('Djarum by Kategori')
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

############ Top 10 Brand Nielsen #############

if option == "Top 10 Brand Nielsen" :

    st.markdown("""---""")

    st.markdown('### Top 10 Brand Nielsen - April 2022 (YTD)')

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

    #############

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

    if st.checkbox ("Show Data (Magnificent 7 Brand Prioritas)"):
        st.markdown('### Magnificent 7 Brand Prioritas')

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

        st.markdown("""---""")    
        st.markdown('### Top 10 Brand Internal - Juni 2022 (YTD)')
        
        names = ['LA Bold', 'Geo Mild', 'Chief Filter','L.A. Lights Regular','D. 7 6', 'D. Super','Chief Kretek']

        df = pd.read_excel(
        io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
        engine="openpyxl",
        sheet_name="DSO",
        ###names=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q'],
        usecols="A:Q",
        nrows=10
        )
        new_df = df.drop(['Group','Kategori', 'Kelas'], axis=1)


    ###    new_df_transpose = new_df.T

        first =((new_df).iloc[0,0])
        first_value = ('{:.2f}%'.format((new_df).iloc[0,9]))
        first_delta = ('{:.2f}%'.format((new_df).iloc[0,13])) 

        second =((new_df).iloc[1,0])
        second_value = ('{:.2f}%'.format((new_df).iloc[1,9]))
        second_delta = ('{:.2f}%'.format((new_df).iloc[1,13]))  

        thirth =((new_df).iloc[2,0])
        thirth_value = ('{:.2f}%'.format((new_df).iloc[2,9]))
        thirth_delta = ('{:.2f}%'.format((new_df).iloc[2,13]))

        fourth =((new_df).iloc[3,0])
        fourth_value = ('{:.2f}%'.format((new_df).iloc[3,9]))
        fourth_delta = ('{:.2f}%'.format((new_df).iloc[3,13]))

        fifth =((new_df).iloc[4,0])
        fifth_value = ('{:.2f}%'.format((new_df).iloc[4,9]))
        fifth_delta = ('{:.2f}%'.format((new_df).iloc[4,13]))

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric(first,first_value, first_delta)
        col2.metric(second,second_value, second_delta)
        col3.metric(thirth, thirth_value, thirth_delta)
        col4.metric(fourth, fourth_value, fourth_delta)
        col5.metric(fifth, fifth_value, fifth_delta)

        #############

        sixth =((new_df).iloc[5,0])
        sixth_value = ('{:.2f}%'.format((new_df).iloc[5,9]))
        sixth_delta = ('{:.2f}%'.format((new_df).iloc[5,13])) 
        seventh =((new_df).iloc[6,0])
        seventh_value = ('{:.2f}%'.format((new_df).iloc[6,9]))
        seventh_delta = ('{:.2f}%'.format((new_df).iloc[6,13]))  
        eighth =((new_df).iloc[7,0])
        eighth_value = ('{:.2f}%'.format((new_df).iloc[7,9]))
        eighth_delta = ('{:.2f}%'.format((new_df).iloc[7,13]))
        ninth =((new_df).iloc[8,0])
        ninth_value = ('{:.2f}%'.format((new_df).iloc[8,9]))
        ninth_delta = ('{:.2f}%'.format((new_df).iloc[8,13]))
        tenth =((new_df).iloc[9,0])
        tenth_value = ('{:.2f}%'.format((new_df).iloc[9,9]))
        tenth_delta = ('{:.2f}%'.format((new_df).iloc[9,13]))

        col6, col7, col8, col9, col10 = st.columns(5)
        col6.metric(sixth,sixth_value, sixth_delta)
        col7.metric(seventh,seventh_value, seventh_delta)
        col8.metric(eighth, eighth_value, eighth_delta)
        col9.metric(ninth, ninth_value, ninth_delta)
        col10.metric(tenth, tenth_value, tenth_delta)

        st.markdown("""---""")
        if st.checkbox ("Show Data (Magnificent 7 Brand Prioritas)"):
            st.markdown('### Magnificent 7 Brand Prioritas')

        ###    st.write(new_df.drop(['Delta', 'Feb21', 'Apr21', 'Agt22', 'Okt22', 'Des22'], axis=1))

            #################
            df2 = pd.read_excel(
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
            engine="openpyxl",
            sheet_name="DSO",
            usecols="A:Q",
            nrows=45
            )

            df_djarum_group = df2[df2.GabBrand.isin(names)]
            df_top_10_internal = df_djarum_group.drop(['Group','Kategori', 'Kelas','Delta', 'Feb21', 'Apr21', 'Agt22', 'Okt22', 'Des22'], axis=1)
            st.write(df_top_10_internal)

            


#####  CODING ANIMATION ######

###def load_lottiefile(filepath: str):
###    with open(filepath, "r") as f:
###     return json.load(f)

###lottie_path_waiting = "/Users/ekasulawestara/Dashboard DSO Pamekasan/waiting.json"
###lottie_loading = load_lottiefile(lottie_path_waiting)

###    with st_lottie_spinner(lottie_loading, height = 500, width = 1000, loop=True, key="loading"):
###        time.sleep(10)


############ DSO Sales Performance #############

if option == "DSO Sales Performance":


    st.markdown("""---""")
    st.subheader(" SALES PERFORMACE")

    @st.cache(allow_output_mutation=True)
    def load_data():
        df = pd.read_excel(
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/Market_DSO.xlsx",
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
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
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
            io="/Users/ekasulawestara/Dashboard DSO Pamekasan/MIT.xlsx",
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

####################
  


if option == "DSO Sales Performance":
    st.markdown("""---""")

    st.markdown('### Kontribusi Omset By Group & Customer Category Sem 1 2022')
    fig = plt.figure(figsize=(12, 5.0625))
    ax1 = fig.add_subplot(121)
    ax2 = fig.add_subplot(122)
    fig.subplots_adjust(wspace=0)
    
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


    if option == "DSO Sales Performance":
        col1, col2 = st.columns(2)
    with col1 :

        fig = px.pie(df, values='OmzetJtBtg', names='CustomerGroupCategory',width= 500, height= 400, title="Kontribusi Omset By Customer Group Category Sem 1 2022")
        
        st.plotly_chart(fig)
    with col2 :

        fig = px.pie(df, values='OmzetJtBtg', names='CustomerCategory',width= 500, height= 400, title="Kontribusi Omset By CustomerCategory Sem 1 2022")
        
        st.plotly_chart(fig)
    



    
