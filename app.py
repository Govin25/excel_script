import streamlit as st
import pandas as pd
import openai
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

# Configure OpenAI
openai.api_key = os.getenv("OPENAI_API_KEY")

def generate_response(question):
    """Generate AI response for a given question"""
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that provides clear and concise answers."},
                {"role": "user", "content": question}
            ],
            max_tokens=150
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating response: {str(e)}"

def process_excel(file):
    """Process Excel file and generate responses"""
    try:
        # Read Excel file
        df = pd.read_excel(file)
        
        # Get the first column name (assuming it contains questions)
        question_column = df.columns[0]
        
        # Generate responses for each question
        df['AI Response'] = df[question_column].apply(generate_response)
        
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
                
                # Download button
                output = df.to_excel(index=False)
                st.download_button(
                    label="Download Results",
                    data=output,
                    file_name="processed_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main() 