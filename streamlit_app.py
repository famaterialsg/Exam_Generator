import streamlit as st
import pandas as pd
import zipfile
import streamlit_ext as ste
import json
import re
import os
from io import BytesIO

# Function to generate exams
def generator(excel_file, number_of_questions):
    temp = []

    for name, question in number_of_questions.items():
        # Read the specific sheet into a DataFrame
        data = pd.read_excel(excel_file, sheet_name=name)

        #if 'Exam NNumber'
        
        # Extract the specified number of random rows from the sheet
        extract = data.sample(question)
        
        # Append the extracted rows to the list
        temp.append(extract)
    
    # Combine all the DataFrames in the list into a single DataFrame
    df_combined = pd.concat(temp, ignore_index=True)

    # Write the combined DataFrame to a new Excel file in memory
    output = BytesIO()
    df_combined.to_excel(output, index=False)
    output.seek(0)
    
    return output, df_combined

# Function for generating exams
def generate_exams():
    st.title("Generate Exams")

    uploaded_file = st.file_uploader("Upload an EXCEL file to get started", type="xlsx")

    if 'generated_files' not in st.session_state:
        st.session_state.generated_files = []

    if uploaded_file is not None:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

        number_of_questions = {}
        for name in sheet_names:
            with st.expander(name):
                selected_number = st.number_input(f"Please enter the number of questions from {name}: ", min_value=0, max_value=10000, step=1)
                number_of_questions[name] = selected_number

        number_of_exams = st.number_input("How many exams do you want to generate?", min_value=1, max_value=100, step=1)
        
        if st.button('Generate'):
            st.session_state.generated_files = []

            for count in range(number_of_exams):
                output_file, df_combined = generator(uploaded_file, number_of_questions)

                json_output = excel_to_json(df_combined)

                st.session_state.generated_files.append((output_file,json_output, df_combined, f"Generated exam number {count + 1}:"))
                
            mem_zip = BytesIO()
            with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for id, (output_file,json_output, df_combined, message) in enumerate(st.session_state.generated_files):
                    with st.expander(message):
                        st.write(df_combined)
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        ste.download_button(
                            label="Download",
                            data=output_file,
                            file_name=f'exam_{id + 1}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            #key=f'download_{id}'
                        )

                    with col2:
                        ste.download_button(
                            label="Download JSON",
                            data=json_output,
                            file_name=f'exam_{id + 1}.json',
                            mime='application/json'
                    )
                    zf.writestr(f'exam_{id + 1}.xlsx', output_file.getvalue())
                    zf.writestr(f'exam_{id + 1}.json', json_output)
        
                mem_zip.seek(0)

            ste.download_button(
                label="Download All as ZIP",
                data=mem_zip.getvalue(),
                file_name='exams.zip',
                mime='application/zip',
                #key="download_all"
            )
            
        #st.write(st.session_state)
    else:
        st.session_state.generated_files = []

def arrange_answers(answers, correct_label):
    correct_index = ord(correct_label.upper()) - ord('A')
    answers.insert(0, answers.pop(correct_index))
    return answers

def clean_text(text):
    text = re.sub(r'<', '&lt;', text)
    text = re.sub(r'>', '&gt;', text)
    text = re.sub(r'\r', '', text)
    text = re.sub(r'\n', '<br>', text)
    return text.strip()

def excel_to_json(data):
    # Prepare the JSON structure
    output_structure = {"mc_questions": []}

    for index, row in data.iterrows():
        try:
            # Extract question and answers
            answers = [row[f'options[{label.lower()}]'] for label in 'ABCDEFG' if pd.notnull(row[f'options[{label.lower()}]'])]
            correct_label = row['correct'].strip().upper()
            # Arrange answers based on the correct label
            arranged_answers = arrange_answers(answers, correct_label) if correct_label in 'ABCDEFG' else answers

            cleaned_question = clean_text(row['question'])
            cleaned_answers = [clean_text(answer) for answer in arranged_answers]

            question_data = {
                "question": cleaned_question,
                "answers": cleaned_answers
            }
            # Add the question data to the list
            output_structure["mc_questions"].append(question_data)
        except KeyError as e:
            print(f"KeyError: {e} at row {index}")
        except Exception as e:
            print(f"Unexpected error: {e} at row {index}")

    # Convert the output structure to a JSON string
    json_data = json.dumps(output_structure, indent=4, ensure_ascii=False)
    
    return json_data


# Function for home page
def home():
    st.title("Welcome to the App")
    st.write("Use the menu to navigate to different functions.")

# Main function to set up the menu and handle navigation
def main():
    st.sidebar.title("Navigation")
    menu_options = ["Home", "Generate Exams"]
    choice = st.sidebar.selectbox("Go to", menu_options)

    if choice == "Home":
        home()
    elif choice == "Generate Exams":
        generate_exams()

if __name__ == "__main__":
    main()
