import json
import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    """테이블 셀의 배경색을 설정합니다."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_borders(cell, color="CBD5E1", sz="4", val="single"):
    """테이블 셀 테두리를 설정합니다."""
    tcPr = cell._element.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="12" w:space="0" w:color="1E3A8A"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

def set_cell_margins(cell, top=140, bottom=140, left=200, right=200):
    """테이블 셀의 내부 패딩을 설정합니다 (dxa 단위: 1pt = 20dxa)."""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_subject_docx(subject_data, output_path, is_combined=False):
    """단일 과목 또는 통합본 Word 문서를 생성합니다."""
    doc = docx.Document()

    # 페이지 여백 설정 (상하좌우 20mm)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)
        
        # 머리글 설정
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("문화창업플래너 자격시험 실전 핵심 기출예상 모의고사")
        hrun.font.name = "맑은 고딕"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(148, 163, 184) # Slate-400

        # 바닥글 설정
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run("- 한국생산성본부 문화창업플래너 자격검정 대비 -")
        frun.font.name = "맑은 고딕"
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(148, 163, 184)

    # 1. 문서 제목
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(4)
    title_p.paragraph_format.space_after = Pt(2)
    title_run = title_p.add_run(subject_data.get("doc_title", subject_data["title"]))
    title_run.font.name = "맑은 고딕"
    title_run.font.size = Pt(17)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138) # Deep Navy (Indigo 900)

    # 부제목
    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(12)
    sub_run = sub_p.add_run("출제위원 직강 녹음 및 강의교재 핵심 포인트 완벽 반영 (기출예상 20선)")
    sub_run.font.name = "맑은 고딕"
    sub_run.font.size = Pt(10)
    sub_run.font.color.rgb = RGBColor(71, 85, 105) # Slate 600

    # 2. 과목 메타정보 안내 박스
    meta_table = doc.add_table(rows=1, cols=1)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_cell = meta_table.cell(0, 0)
    meta_cell.width = Inches(6.8)
    set_cell_background(meta_cell, "F1F5F9") # Slate 100
    set_cell_borders(meta_cell, color="CBD5E1", sz="6", val="single")
    set_cell_margins(meta_cell, top=140, bottom=140, left=180, right=180)

    mp = meta_cell.paragraphs[0]
    mp.paragraph_format.line_spacing = 1.25
    mp.paragraph_format.space_after = Pt(2)

    def add_meta_line(p, label, val):
        r_lbl = p.add_run(f"• {label}: ")
        r_lbl.font.name = "맑은 고딕"
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = RGBColor(30, 41, 59)

        r_val = p.add_run(f"{val}\n")
        r_val.font.name = "맑은 고딕"
        r_val.font.size = Pt(9.5)
        r_val.font.color.rgb = RGBColor(71, 85, 105)

    add_meta_line(mp, "출제 및 강의", subject_data.get("instructor", "출제위원"))
    add_meta_line(mp, "합격 기준", "과목당 40점 이상(40점 미만 과락 탈락), 전과목 평균 60점 이상 합격")
    add_meta_line(mp, "출제 핵심 포인트", subject_data.get("description", ""))

    # 메타박스 끝 줄바꿈 제거
    mp.runs[-1].text = mp.runs[-1].text.rstrip("\n")

    # 빈 줄 추가
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(8)
    p_div.paragraph_format.space_after = Pt(4)

    # 3. 문항 리스트 출력
    opt_labels = ["①", "②", "③", "④"]
    questions = subject_data["questions"]

    for idx, q in enumerate(questions, 1):
        # [문항 N] 번호 및 문제 본문
        qp = doc.add_paragraph()
        qp.paragraph_format.space_before = Pt(14)
        qp.paragraph_format.space_after = Pt(6)
        qp.paragraph_format.line_spacing = 1.25

        q_num_run = qp.add_run(f"[문항 {idx:02d}] ")
        q_num_run.font.name = "맑은 고딕"
        q_num_run.font.size = Pt(11)
        q_num_run.font.bold = True
        q_num_run.font.color.rgb = RGBColor(30, 58, 138) # Deep Navy

        q_text_run = qp.add_run(q["question"])
        q_text_run.font.name = "맑은 고딕"
        q_text_run.font.size = Pt(10.5)
        q_text_run.font.bold = True
        q_text_run.font.color.rgb = RGBColor(15, 23, 42) # Slate 900

        # 선택지 (4지선다)
        for opt_idx, opt_text in enumerate(q["options"]):
            opt_p = doc.add_paragraph()
            opt_p.paragraph_format.left_indent = Inches(0.25)
            opt_p.paragraph_format.space_before = Pt(1.5)
            opt_p.paragraph_format.space_after = Pt(1.5)
            opt_p.paragraph_format.line_spacing = 1.2

            lbl_run = opt_p.add_run(f"{opt_labels[opt_idx]} ")
            lbl_run.font.name = "맑은 고딕"
            lbl_run.font.size = Pt(9.5)
            lbl_run.font.bold = True
            lbl_run.font.color.rgb = RGBColor(71, 85, 105)

            txt_run = opt_p.add_run(opt_text)
            txt_run.font.name = "맑은 고딕"
            txt_run.font.size = Pt(9.5)
            txt_run.font.color.rgb = RGBColor(51, 65, 85)

        # 해설 및 힌트 박스 (Callout 테이블)
        sol_table = doc.add_table(rows=1, cols=1)
        sol_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sol_cell = sol_table.cell(0, 0)
        sol_cell.width = Inches(6.8)
        set_cell_background(sol_cell, "F8FAFC") # Slate 50
        set_cell_borders(sol_cell, color="E2E8F0", sz="4", val="single")
        set_cell_margins(sol_cell, top=100, bottom=100, left=150, right=150)

        sp = sol_cell.paragraphs[0]
        sp.paragraph_format.line_spacing = 1.25
        sp.paragraph_format.space_after = Pt(2)

        # 정답
        ans_lbl = sp.add_run(f"【정답】 {opt_labels[q['answer']]}번\n")
        ans_lbl.font.name = "맑은 고딕"
        ans_lbl.font.size = Pt(9.5)
        ans_lbl.font.bold = True
        ans_lbl.font.color.rgb = RGBColor(5, 150, 105) # Emerald 600

        # 강사 출제 힌트
        if q.get("hint"):
            hint_lbl = sp.add_run("【강사 출제 힌트】 ")
            hint_lbl.font.name = "맑은 고딕"
            hint_lbl.font.size = Pt(9)
            hint_lbl.font.bold = True
            hint_lbl.font.color.rgb = RGBColor(217, 119, 6) # Amber 600

            hint_txt = sp.add_run(f"{q['hint']}\n")
            hint_txt.font.name = "맑은 고딕"
            hint_txt.font.size = Pt(9)
            hint_txt.font.color.rgb = RGBColor(146, 64, 14) # Amber 800

        # 상세 해설
        exp_lbl = sp.add_run("【상세 해설】 ")
        exp_lbl.font.name = "맑은 고딕"
        exp_lbl.font.size = Pt(9)
        exp_lbl.font.bold = True
        exp_lbl.font.color.rgb = RGBColor(30, 41, 59)

        exp_txt = sp.add_run(q["explanation"])
        exp_txt.font.name = "맑은 고딕"
        exp_txt.font.size = Pt(9)
        exp_txt.font.color.rgb = RGBColor(71, 85, 105)

    doc.save(output_path)
    print(f"Generated: {output_path}")

def main():
    # data.json 로드
    with open("data.json", "r", encoding="utf-8") as f:
        quiz_data = json.load(f)

    # 출력 폴더 1: 로컬 프로젝트 폴더
    local_dir = r"C:\Users\taeni\.gemini\antigravity\scratch\cultural_startup_planner_cbt"
    # 출력 폴더 2: 구글 드라이브 동기화 폴더
    gdrive_dir = r"g:\내 드라이브\창업관련\문화창업플래너 자격시험 자료\2회"

    # 파일명 매핑
    subject_filenames = {
        "subject1": "01_과목1_문화콘텐츠_창업의_이해_예상기출문제_20선.docx",
        "subject2": "02_과목2_문화콘텐츠_창업기획_및_전략수립_예상기출문제_20선.docx",
        "subject3": "03_과목3_문화콘텐츠기업_자금조달_예상기출문제_20선.docx",
        "subject4": "04_과목4_문화콘텐츠기업_조직관리_예상기출문제_20선.docx"
    }

    generated_files = []

    # 1. 과목별 개별 문서 생성
    for s_key, filename in subject_filenames.items():
        s_data = quiz_data[s_key]
        s_data["doc_title"] = f"문화창업플래너 자격시험: {s_data['title']}"
        local_path = os.path.join(local_dir, filename)
        create_subject_docx(s_data, local_path)
        generated_files.append((local_path, filename))

    # 2. 전과목 종합 실전 모의고사 (80선) 생성
    all_questions = []
    for s_key in ["subject1", "subject2", "subject3", "subject4"]:
        for q in quiz_data[s_key]["questions"]:
            q_copy = dict(q)
            q_copy["question"] = f"[{quiz_data[s_key]['title'].split(':')[0]}] " + q_copy["question"]
            all_questions.append(q_copy)

    combined_data = {
        "title": "전과목 종합 실전 모의고사 (80문항)",
        "doc_title": "문화창업플래너 자격시험 실전 종합 모의고사 80선",
        "instructor": "백상훈 대표, 이일한 교수, 오상민 대표, 김은경 위촉교수",
        "description": "1과목(창업이해) + 2과목(기획·전략) + 3과목(자금조달) + 4과목(조직관리) 전과목 80문항 총괄 테스트 (시험시간 100분)",
        "questions": all_questions
    }
    combined_filename = "문화창업플래너_전과목_종합실전모의고사_80선.docx"
    combined_path = os.path.join(local_dir, combined_filename)
    create_subject_docx(combined_data, combined_path, is_combined=True)
    generated_files.append((combined_path, combined_filename))

    # 3. 구글 드라이브 동기화 폴더로 복사
    if os.path.exists(gdrive_dir):
        print(f"\nCopying files to Google Drive directory: {gdrive_dir}")
        for l_path, f_name in generated_files:
            g_path = os.path.join(gdrive_dir, f_name)
            shutil.copy2(l_path, g_path)
            print(f"Copied to Google Drive: {g_path}")

    print("\nAll docx files generated successfully!")

if __name__ == "__main__":
    main()
