from docx import Document
from docx.shared import Pt

# 템플릿 문서 경로
template_path = '사용자 가이드 - RAG(서울).docx'
passwords_list_file = 'pw-rag_seoul.txt'

def copy_and_modify_template(template_path, output_path, replacement, i):
    # 기존 문서를 로드
    doc = Document(template_path)
    
    # "password"와 "0000"를 대체할 내용으로 수정
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if "password" in run.text:
                run.text = run.text.replace("password", replacement)
                run.font.size = Pt(12)  # 원래 font-size 유지
            elif "0000" in paragraph.text:
                # 대체할 문자열 생성
                replacement_text = f"{'0' * (4 - len(str(i+1)))}{i+1}"
                if "0000" in paragraph.text:
                    paragraph.text = paragraph.text.replace("0000", replacement_text)
                    for run_0000 in paragraph.runs:
                        if "0000" in run_0000.text:
                            run_0000.text = run_0000.text.replace("0000", replacement_text)
                        run_0000.font.size = Pt(12)  # 원래 font-size 유지

                        # "인지게임/사용자 정보 페이지 아이디:" 부분을 bold로 설정
                        if "인지게임/사용자 정보 페이지 아이디:" in run_0000.text:
                            run_0000.bold = True
                            break

    doc.save(output_path)
    print(f"Saving file: {output_path}")


with open(passwords_list_file, 'r') as file:
    replacement_list = file.read().splitlines()

# "password"와 "0000"를 대체하는 반복문
for i, replacement in enumerate(replacement_list):
    output_path = f'document_{i+1}.docx'
    copy_and_modify_template(template_path, output_path, replacement, i)
