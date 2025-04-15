import streamlit as st
import pandas as pd
from datetime import datetime

# Define CSV file location
csv_file = "scoreboard.csv"

# Load existing data
try:
    df = pd.read_csv(csv_file)
except FileNotFoundError:
    df = pd.DataFrame(columns=["Date", "Dragon Warrior", "Iron Phantom", "Score"])

# Ensure correct column names & strip any extra spaces
df.columns = df.columns.str.strip()

# Show current scoreboard
st.title("🏆 Pool Scoreboard App")
st.write("Daily match scores & overall leaderboard")

# Display scores in "Dragon Warrior - Iron Phantom" format
if "Dragon Warrior" in df.columns and "Iron Phantom" in df.columns:
    df["Score"] = df["Dragon Warrior"].astype(str) + "-" + df["Iron Phantom"].astype(str)
else:
    st.error("Error: Column names in scoreboard.csv are incorrect! Please check headers.")

st.dataframe(df[["Date", "Score"]])

# Input new scores
st.subheader("📅 Add Today's Scores")
date_today = datetime.today().strftime('%Y-%m-%d')
dragon_warrior_score = st.number_input("Dragon Warrior Score", min_value=0)
iron_phantom_score = st.number_input("Iron Phantom Score", min_value=0)
submit = st.button("Save Scores")

# Update CSV and save new scores
if submit:
    new_data = {"Date": date_today, "Dragon Warrior": dragon_warrior_score, "Iron Phantom": iron_phantom_score, "Score": f"{dragon_warrior_score}-{iron_phantom_score}"}
    
    # Use pd.concat instead of deprecated .append()
    df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
    df.to_csv(csv_file, index=False)
    
    st.success(f"Scores saved: {dragon_warrior_score}-{iron_phantom_score} for {date_today} 🎱")
    st.balloons()

# Calculate total scores
total_scores = df[["Dragon Warrior", "Iron Phantom"]].sum()
st.subheader("🏅 Total Scores")
st.write(f"Dragon Warrior: **{total_scores['Dragon Warrior']}** | Iron Phantom: **{total_scores['Iron Phantom']}**")

# Determine who is winning
winning_player = "Dragon Warrior" if total_scores["Dragon Warrior"] > total_scores["Iron Phantom"] else "Iron Phantom"
st.subheader(f"🥇 Leading Player: {winning_player}!")
