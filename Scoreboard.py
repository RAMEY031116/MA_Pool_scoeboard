import streamlit as st
import pandas as pd
from datetime import datetime

csv_file = "pool_scores.csv"

# Load existing data
try:
    df = pd.read_csv(csv_file)
except FileNotFoundError:
    df = pd.DataFrame(columns=["Date", "Player 1", "Player 2", "Score"])

# Show current scoreboard
st.title("🏆 Pool Scoreboard App")
st.write("Daily match scores & overall leaderboard")

# Format scores as "Player 1 - Player 2"
df["Score"] = df["Player 1"].astype(str) + "-" + df["Player 2"].astype(str)
st.dataframe(df[["Date", "Score"]])

# Input new scores
st.subheader("📅 Add Today's Scores")
date_today = datetime.today().strftime('%Y-%m-%d')
player1_score = st.number_input("Player 1 Score", min_value=0)
player2_score = st.number_input("Player 2 Score", min_value=0)
submit = st.button("Save Scores")

# Update CSV
if submit:
    new_data = {"Date": date_today, "Player 1": player1_score, "Player 2": player2_score, "Score": f"{player1_score}-{player2_score}"}
    df = df.append(new_data, ignore_index=True)
    df.to_csv(csv_file, index=False)
    
    st.success(f"Scores saved: {player1_score}-{player2_score} for {date_today} 🎱")
    st.balloons()

# Calculate total scores
total_scores = df[["Player 1", "Player 2"]].sum()
st.subheader("🏅 Total Scores")
st.write(f"Player 1: **{total_scores['Player 1']}** | Player 2: **{total_scores['Player 2']}**")

# Determine who's winning
winning_player = "Player 1" if total_scores["Player 1"] > total_scores["Player 2"] else "Player 2"
st.subheader(f"🥇 Leading Player: {winning_player}!")
