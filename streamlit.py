import streamlit as st
import pandas as pd
# import pydeck as pdk
from bs4 import BeautifulSoup
import requests
import csv
import json
import re
import asyncio
# from operator import itemgetter
# from urllib.error import URLError
# from streamlit_observable import observable
# import pandas_profiling
# from streamlit_pandas_profiling import st_profile_report
import numpy as np
from property import property
from prologis import export

st.set_page_config(page_title="Prologis Fulcrum Export", layout="wide")

# PAGES = {
#     "Home": intro,
#     "Your Favorite Actors": actors,
#     "Your Favorite Directors": director,
#     "Your Favorite Movies By Length": length,
#     "Your Favorite Genres": genre,
#     "Your Favorite Decades": decade,
#     "Your Favorite Languages": language,
#     # "Your Favorite Country": country,
#     # "Your Favorite Year": year,
#     "Recommendations All": recommender,
#     "Recommendations": recommender2
# }

st.header('Prologis Fulcrum Export')
# st.write('TO PUT HERE.....')
options = ['NNJ00116', 'NNJ09304', 'NNJ00117',
            'NNJ06904', 'G-IIG231', 'NNJ09308', 'NNJ07503', 'NNJ06304']

property_code = st.selectbox('Which property do you want to export?', options)
# st.button('Choose Property Code', on_click=property, args=(property_code, ))
# print(option)
# Define a default value for the session variable
# if "my_global_variable" not in st.session_state:
#     st.session_state.my_global_variable = options[0]

st.write('You selected:', property_code)


# st.write('Export:', st.session_state.my_global_variable)
st.button('Export Property Code', on_click=export, args=(property_code, ))

# df = pd.read_csv(file)
# df = df.dropna(subset=['MyRating'])
# dfAvg = df[(df["MovieLength"] > 60) & (df["MovieLength"] < 275)]
# avg = dfAvg["MyRating"].sum() / dfAvg["Movie"].count()

# df["Genre"] = df["Genre"].str.split(",")
# df["Languages"] = df["Languages"].str.split(",")
# df["Actors"] = df["Actors"].str.split(",")

# df = df.drop(["MovieLength", "NumberOfReviews"], axis=1)
# df.index = df.index + 1
# df = df.rename(columns={"MyRating": "Your Rating", "LBRating": "Letterboxd Rating", "ReviewDate": "Date Reviewed",
#                 "LengthInHour": "Movie Length", "Genre": "Genres", "NumberOfRatings": "Number Of Ratings", "ReleaseYear": "Release Year"})
# pd.options.mode.chained_assignment = None
# # df2 = df.style.background_gradient(subset=['Ranking', 'Billing Score'])

# avgRound = "{:.2f}".format(avg)
# totalMovies = len(df)
# st.write(
#     f'Your average rating across all movies is: **{avgRound}** over **{totalMovies}** amount of movies')

# st.dataframe(df, height=700, use_container_width=True)