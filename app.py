import streamlit as st
import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
import os
from io import BytesIO
import time
from datetime import datetime
import asyncio
import aiohttp
import backoff
from concurrent.futures import ThreadPoolExecutor
from typing import List, Tuple, Dict, Optional
import json

# Load environment variables
load_dotenv()

# Configure OpenAI
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Constants
MAX_RETRIES = 3
BATCH_SIZE = 5  # Number of questions to process in parallel
MAX_WORKERS = 3  # Number of worker threads

@backoff.on_exception(
    backoff.expo,
    Exception,
    max_tries=MAX_RETRIES,
    giveup=lambda e: "insufficient_quota" in str(e)
)
async def generate_response_with_backoff(question: str) -> Tuple[str, float]:
    """Generate AI response with exponential backoff retry"""
    start_time = time.time()
    try:
        response = await asyncio.to_thread(
            client.chat.completions.create,
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that provides clear and concise answers."},
                {"role": "user", "content": question}
            ],
            max_tokens=150
        )
        end_time = time.time()
        return response.choices[0].message.content, end_time - start_time
    except Exception as e:
        error_message = str(e)
        if "insufficient_quota" in error_message:
            return "Error: OpenAI API quota exceeded. Please check your billing status.", 0
        raise  # Re-raise the exception for backoff to handle

async def process_batch(questions: List[str]) -> List[Tuple[str, float]]:
    """Process a batch of questions concurrently"""
    tasks = [generate_response_with_backoff(q) for q in questions]
    return await asyncio.gather(*tasks, return_exceptions=True)

def create_progress_html(current: int, total: int, batch_size: int) -> str:
    """Create HTML for progress display"""
    progress = current / total
    return f"""
        <div style="display: flex; align-items: center; gap: 10px;">
            <div style="flex-grow: 1;">
                <div class="stProgress">
                    <div style="background-color: #e0e0e0; height: 20px; border-radius: 10px; overflow: hidden;">
                        <div style="background-color: #1f77b4; width: {progress*100}%; height: 100%;"></div>
                    </div>
                </div>
            </div>
            <div style="white-space: nowrap;">
                Processing questions {current+1}-{min(current+batch_size, total)} of {total}
            </div>
        </div>
    """

def update_progress(placeholder: st.empty, current: int, total: int, batch_size: int) -> None:
    """Update the progress display"""
    progress_html = create_progress_html(current, total, batch_size)
    placeholder.markdown(progress_html, unsafe_allow_html=True)

def process_batch_results(batch_results: List[Tuple[str, float]]) -> Tuple[List[str], List[float]]:
    """Process results from a batch of API calls"""
    processed_questions = []
    processed_times = []
    
    for response, processing_time in batch_results:
        if isinstance(response, Exception):
            response = f"Error: {str(response)}"
        processed_questions.append(response)
        processed_times.append(processing_time)
    
    return processed_questions, processed_times

def display_processing_stats(total_questions: int, total_time: float, processed_times: List[float]) -> None:
    """Display processing statistics"""
    st.write("### Processing Statistics")
    st.write(f"Total number of questions: {total_questions}")
    st.write(f"Total processing time: {total_time:.2f} seconds")
    st.write(f"Average time per question: {(total_time/total_questions):.2f} seconds")
    st.write(f"Maximum processing time: {max(processed_times):.2f} seconds")
    st.write(f"Minimum processing time: {min(processed_times):.2f} seconds")

def process_excel(file) -> Optional[pd.DataFrame]:
    """Process Excel file and generate responses"""
    try:
        # Read Excel file
        df = pd.read_excel(file)
        question_column = df.columns[0]
        
        # Initialize timing metrics
        total_start_time = time.time()
        
        # Create a placeholder for progress and counter
        progress_placeholder = st.empty()
        
        # Process questions in batches
        questions = df[question_column].tolist()
        total_questions = len(questions)
        processed_questions = []
        processed_times = []
        
        for i in range(0, total_questions, BATCH_SIZE):
            batch = questions[i:i + BATCH_SIZE]
            
            # Process batch concurrently
            batch_results = asyncio.run(process_batch(batch))
            
            # Handle results
            batch_questions, batch_times = process_batch_results(batch_results)
            processed_questions.extend(batch_questions)
            processed_times.extend(batch_times)
            
            # Update progress
            update_progress(progress_placeholder, i, total_questions, BATCH_SIZE)
        
        # Update DataFrame with results
        df['Answers'] = processed_questions
        df['Processing Time (seconds)'] = processed_times
        
        total_end_time = time.time()
        total_time = total_end_time - total_start_time
        
        # Display statistics
        display_processing_stats(total_questions, total_time, processed_times)
        
        return df
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
        return None

def main():
    st.title("Excel Question Answerer")
    st.write("Upload an Excel file with questions in the first column to get AI-generated responses.")
    
    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Process the file
        with st.spinner("Processing your Excel file..."):
            df = process_excel(uploaded_file)
            
            if df is not None:
                # Display the results
                st.write("### Preview of Results")
                st.dataframe(df)
                
                # Create a buffer to store the Excel file
                buffer = BytesIO()
                df.to_excel(buffer, index=False)
                buffer.seek(0)
                
                # Download button
                st.download_button(
                    label="Download Results",
                    data=buffer,
                    file_name="processed_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main() 