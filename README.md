# Excel Question Answerer

This application allows you to upload an Excel file containing questions and generates AI-powered responses for each question. The responses are added in a new column and can be downloaded as an Excel file.

## Setup

1. Create and activate a virtual environment (Windows):
```bash
# Go to root directory and Create a virtual environment
python -m venv venv

# Activate the virtual environment
.\venv\Scripts\activate
```

2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root and add your OpenAI API key:
```
OPENAI_API_KEY=your_api_key_here
```

## Usage

1. Make sure your virtual environment is activated (you should see `(venv)` in your terminal prompt)

2. Run the application:
```bash
streamlit run app.py
```

3. Open your web browser and navigate to the URL shown in the terminal (typically http://localhost:8501)

4. Upload an Excel file (.xlsx or .xls) with questions in the first column

5. Wait for the processing to complete

6. Preview the results and download the processed Excel file

## Excel File Format

- The first column should contain your questions
- The application will automatically add a new column named "AI Response" with the generated answers
- The number of rows can be dynamic

## Notes

- Make sure you have a valid OpenAI API key
- The application uses GPT-3.5-turbo model for generating responses
- Processing time depends on the number of questions in your Excel file
- To deactivate the virtual environment when you're done, simply type `deactivate` in the terminal 
