import pyodbc
import pandas as pd
import streamlit as st


@st.cache_resource
def get_connection():
    server   = st.secrets["db"]["server"]
    database = st.secrets["db"]["database"]
    username = st.secrets["db"]["username"]
    password = st.secrets["db"]["password"]
    conn_str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={server};DATABASE={database};UID={username};PWD={password};"
    )
    return pyodbc.connect(conn_str)


@st.cache_data(ttl=3600)
def load_data() -> pd.DataFrame:
    conn = get_connection()

    df_customer = pd.read_sql("""
        SELECT T0.CustomerID, T0.CustomerName, T0.Address, T0.RegionID,
               T0.CityID, T0.AccountID, T1.AccountName,T1.AccountDisplay,
               CASE
                   WHEN T0.ChannelName = '21-Gov. Hosp - Dept'    THEN '101'
                   WHEN T0.ChannelName = '22-Gov. Hosp - Phar'    THEN '102'
                   WHEN T0.ChannelName = '23-Private Hosp - Dept' THEN '104'
                   WHEN T0.ChannelName = '24-Private Hosp - Phar' THEN '106'
                   WHEN T0.ChannelName = '25-Polyclinic'          THEN '103'
                   WHEN T0.ChannelName = '26-Pharmacy'            THEN '105'
                   WHEN T0.ChannelName = '27-Wholesalers'         THEN '107'
                   WHEN T0.ChannelName = '28-Monoclinic'          THEN '108'
                   WHEN T0.ChannelName = '29-Pharmacy Chain'      THEN '109'
               END AS ChannelID
        FROM [VNDB_O365].[dbo].[vw_Sipro_Customer] AS T0
        LEFT JOIN [VNDB_O365].[dbo].[vw_ParentAccount] T1
            ON T0.CUSTOMERID = T1.CUSTOMERID
    """, conn)

    df_sales = pd.read_sql("""
        SELECT T0.CustomerID, T0.Qty, T1.TerritoryID, T0.Cust_Brand,
               T0.BrandID, T0.ProductID, T0.ByMonth, T0.ByYear,
               T0.YYMM, T0.AmountUSD, T1.ProductWeight
        FROM (
            SELECT CustomerID, AmountUSD, Qty, BrandID,
                   ProductID, ByMonth, ByYear,
                   YYMM, CONCAT(CustomerID, Brand) AS Cust_Brand
            FROM [VNDB_O365].[dbo].[vw_SIpro_CustomerProductByCutoff]
        ) AS T0
        LEFT JOIN (
            SELECT TeamID, TerritoryID, CustomerID, BrandID,
                   ProductWeight, Brand,
                   CONCAT(CustomerID, Brand) AS Cust_Brand
            FROM [VNDB_O365].[dbo].[vw_SIpro_TerritoryAlignment]
        ) AS T1 ON T0.Cust_Brand = T1.Cust_Brand
    """, conn)

    df_brand = pd.read_excel(r"C:\Users\khwk503\AZCollaboration\AZVN ComEx - AZVN Dashboard\Pre_CP_Monthly_Dashboard\CategoryBank.xlsx", sheet_name="Brand")
    df_region = pd.read_excel(r"C:\Users\khwk503\AZCollaboration\AZVN ComEx - AZVN Dashboard\Pre_CP_Monthly_Dashboard\CategoryBank.xlsx", sheet_name="Region")
    df_channel = pd.read_excel(r"C:\Users\khwk503\AZCollaboration\AZVN ComEx - AZVN Dashboard\Pre_CP_Monthly_Dashboard\CategoryBank.xlsx", sheet_name="Channel")    




    df = df_sales.merge(df_customer, on="CustomerID", how="left")

    #df = df.merge(df_brand[["BrandID", "BrandCategory"]], on="BrandID", how="left")
    return df
