import streamlit as st
import pandas as pd
import sqlite3 
conn = sqlite3.connect('database.db',check_same_thread=False)
c = conn.cursor()

def create_projectable():
      c.execute('CREATE TABLE IF NOT EXISTS historique(project_name TEXT,num_project TEXT,phase TEXT , reseau TEXT,ligne TEXT,lieu TEXT,version TEXT, d DATE)')
        
def add_projectdata(project_name,num_project,phase,reseau,ligne,lieu,version,date):
    c.execute('INSERT INTO historique(project_name ,num_project ,phase , reseau ,ligne,lieu,version,d) VALUES (?,?,?,?,?,?,?,?)',(project_name,num_project,phase,reseau,ligne,lieu,version,date))
    conn.commit()
    
    
def view_all_projects():
    c.execute('SELECT * FROM historique')
    data = c.fetchall()
    return data