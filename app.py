import streamlit as st
import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
import os
from io import BytesIO
import time
from datetime import datetime

# Load environment variables
load_dotenv()

# Configure OpenAI
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def generate_response(question):
    """Generate AI response for a given question"""
    try:
        start_time = time.time()
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that provides clear and concise answers."},
                {"role": "user", "content": question}
            ],
            max_tokens=150
        )
        end_time = time.time()
        processing_time = end_time - start_time
        return response.choices[0].message.content, processing_time
    except Exception as e:
        error_message = str(e)
        if "insufficient_quota" in error_message:
            return "Error: OpenAI API quota exceeded. Please check your billing status and plan details at https://platform.openai.com/account/billing/overview", 0
        return f"Error generating response: {error_message}", 0

def process_excel(file):
    """Process Excel file and generate responses"""
    try:
        # Read Excel file
        df = pd.read_excel(file)
        
        # Get the first column name (assuming it contains questions)
        question_column = df.columns[0]
        
        # Initialize timing metrics
        total_start_time = time.time()
        processing_times = []
        
        # Create a placeholder for progress and counter
        progress_placeholder = st.empty()
        
        # Generate responses for each question
        for idx, question in enumerate(df[question_column]):
            response, processing_time = generate_response(question)
            df.at[idx, 'Answers'] = response
            processing_times.append(processing_time)
            
            # Update progress and counter in the same line
            progress = (idx + 1) / len(df)
            progress_html = f"""
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="flex-grow: 1;">
                        <div class="stProgress">
                            <div style="background-color: #e0e0e0; height: 20px; border-radius: 10px; overflow: hidden;">
                                <div style="background-color: #1f77b4; width: {progress*100}%; height: 100%;"></div>
                            </div>
                        </div>
                    </div>
                    <div style="white-space: nowrap;">
                        Processing question {idx + 1} of {len(df)}
                    </div>
                </div>
            """
            progress_placeholder.markdown(progress_html, unsafe_allow_html=True)
        
        total_end_time = time.time()
        total_time = total_end_time - total_start_time
        
        # Add timing information
        df['Processing Time (seconds)'] = processing_times

        # Log timing information
        st.write("### Processing Statistics")
        st.write(f"Total number of questions: {len(df)}")
        st.write(f"Total processing time: {total_time:.2f} seconds")
        st.write(f"Average time per question: {(total_time/len(df)):.2f} seconds")
        st.write(f"Maximum processing time: {max(processing_times):.2f} seconds")
        st.write(f"Minimum processing time: {min(processing_times):.2f} seconds")
        
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