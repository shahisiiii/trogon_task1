import json
from docx import Document

def parse_questions(file_path):
    try:
        doc = Document(file_path)
        questions = []
        current_question = None

        for paragraph in doc.paragraphs:
            stripped_line = paragraph.text.strip()
            if stripped_line.startswith("Answer:"):
                current_question['answer'] = stripped_line.split(":")[1].strip()
                questions.append(current_question)
                current_question = None
            elif stripped_line:
                if stripped_line[0].isdigit():
                    if current_question:
                        questions.append(current_question)
                    current_question = {'question': stripped_line}
                elif current_question:
                    if 'options' not in current_question:
                        current_question['options'] = []
                    current_question['options'].append(stripped_line)

        if current_question:
            questions.append(current_question)

        return questions
    except FileNotFoundError:
        print("File not found.")
        return None
    except Exception as e:
        print("An error occurred:", e)
        return None

def convert_to_json(questions, output_path):
    if questions:
        with open(output_path, 'w') as json_file:
            json.dump(questions, json_file, indent=4)
        print("Conversion to JSON successful.")
    else:
        print("No questions found.")

if __name__ == "__main__":
    input_file_path = "NewQuestionAndAnswers.docx"
    output_file_path = "parsed_questions.json"

    parsed_questions = parse_questions(input_file_path)
    convert_to_json(parsed_questions, output_file_path)
