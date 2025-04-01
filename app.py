import streamlit as st
import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
import os
from io import BytesIO

# Load environment variables
load_dotenv()

# Configure OpenAI
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def generate_response(question):
    """Generate AI response for a given question"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that provides clear and concise answers."},
                {"role": "user", "content": question}
            ],
            max_tokens=150
        )
        return response.choices[0].message.content
    except Exception as e:
        error_message = str(e)
        if "insufficient_quota" in error_message:
            return "Error: OpenAI API quota exceeded. Please check your billing status and plan details at https://platform.openai.com/account/billing/overview"
        return f"Error generating response: {error_message}"

def process_excel(file):
    """Process Excel file and generate responses"""
    try:
        # Read Excel file
        df = pd.read_excel(file)
        
        # Get the first column name (assuming it contains questions)
        question_column = df.columns[0]
        
        # Generate responses for each question
        df['Answers'] = df[question_column].apply(generate_response)
        
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