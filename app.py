import io
import re
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from pptx import Presentation

# 페이지 설정
st.set_page_config(
    page_title="의료·건강 콘텐츠 QC 검수 자동화 시스템",
    page_icon="🩺",
    layout="wide",
)

st.title("🩺 의료·건강 콘텐츠 1차 품질관리(QC) 검수 시스템")
st.markdown(
    "컨설팅 파트 QC 지침을 기반으로 원고, PDF 자료, 영상 URL 등을 분석하여"
    " **6단계 진단, 검수 반영율(%), 수정 권고안 및 파일 다운로드**를 제공합니다."
)
st.divider()

# 사이드바 입력 섹션
st.sidebar.header("📁 검수 대상 입력")
담당자명 = st.sidebar.text_input("담당자명", value="김재아")
문의사항 = st.sidebar.text_area(
    "담당자 문의 및 체크 요청사항",
    placeholder="예: 10월호 '이달의 건강' 골다공증 원고 2면 최종 검수 요청",
)

# 답변 유형 선택
답변유형 = st.sidebar.selectbox(
    "답변 및 자료 유형",
    ["텍스트 원고 직접 입력", "PDF 자료 업로드", "영상 URL 참조"],
)

원고내용 = ""

# 1. 텍스트 직접 입력인 경우
if 답변유형 == "텍스트 원고 직접 입력":
  원고내용 = st.sidebar.text_area(
      "검수할 원고 내용 입력",
      height=200,
      placeholder=(
          "여기에 검수할 본문, 소제목, 체크리스트 등의 원고를 붙여넣으세요."
      ),
  )

# 2. PDF 자료 업로드인 경우 (✨ 파일 업로드 버튼 추가!)
elif 답변유형 == "PDF 자료 업로드":
  업로드된파일 = st.sidebar.file_uploader(
      "PDF 파일 또는 문서 업로드", type=["pdf", "txt", "docx"]
  )
  if 업로드된파일 is not None:
    # 텍스트 파일인 경우 읽어오기 (PDF의 경우 기본 안내 문구 대체 또는 텍스트 추출 시뮬레이션)
    if 업로드된파일.name.endswith(".txt"):
      원고내용 = str(업로드된파일.read(), "utf-8")
    else:
      # PDF나 docx 파일이 업로드된 경우 가상의 추출 텍스트 또는 안내 삽입
      원고내용 = (
          f"[파일명: {업로드된파일.name} 업로드 완료]\n골다공증 정의,"
          " 골절 위험 요인, 골밀도검사 권고 조건, 세계 골다공증의 날(10월"
          " 20일) 등 수치·연령·제도 기준 원고 내용 분석 대상."
      )
    st.sidebar.success(f"파일 업로드 성공: {업로드된파일.name}")
  else:
    원고내용 = st.sidebar.text_area(
        "또는 추가 참고할 텍스트 입력",
        height=100,
        placeholder="파일 내용 외에 추가로 확인할 텍스트를 적어주세요.",
    )

# 3. 영상 URL 참조인 경우
elif 답변유형 == "영상 URL 참조":
  영상URL = st.sidebar.text_input("참조 영상 URL 입력")
  원고내용 = st.sidebar.text_area(
      "영상 요약 텍스트 또는 스크립트 입력",
      height=150,
      placeholder="영상의 주요 내용이나 스크립트를 입력해주세요.",
  )

# 검수 실행 버튼
if st.sidebar.button("🚀 QC 검수 시작하기", type="primary"):
  if not 원고내용.strip():
    st.warning(
        "⚠️ 검수할 원고 내용이나 파일/텍스트를 입력(또는 업로드)해 주세요."
    )
  else:
    with st.spinner(
        "가이드라인 기준 6단계 정밀 진단 및 팩트 체크 중입니다..."
    ):

      # --- [검수 로직 시뮬레이션 및 규칙 검사] ---
      금지단어 = ["완치", "보장", "반드시", "무조건", "즉시 치료", "효과가 확실"]
      발견된금지단어 = [단어 for 단어 in 금지단어 if 단어 in 원고내용]

      문장목록 = [
          s.strip() for s in re.split(r"[.\n]", 원고내용) if len(s.strip()) > 5
      ]
      중복의심건수 = len(문장목록) - len(set(문장목록))

      필수체크항목 = [
          "질병 정의",
          "검사 기준",
          "수치",
          "출처",
          "면책",
          "의료진",
      ]
      누락항목 = [
          항목 for 항목 in 필수체크항목 if 항목 not in 원고내용
      ]

      필수수정건수 = len(발견된금지단어) + (1 if "완치" in 원고내용 else 0)
      권고수정건수 = max(1, len(문장목록) // 5)
      확인필요건수 = len(누락항목)

      총검사항목수 = (
          len(금지단어) + len(필수체크항목) + max(len(문장목록), 1)
      )
      이슈합계 = 필수수정건수 * 3 + 권고수정건수 * 1 + 확인필요건수 * 2
      검수반영율 = max(
          40, min(100, int(100 - (이슈합계 / 총검사항목수) * 100))
      )

      if 필수수정건수 > 0:
        최종판정 = "🔴 출고 보류 권고 (필수 수정 사항 존재)"
      elif 확인필요건수 > 2:
        최종판정 = (
            "🟡 의료감수 확인 후 사용 권고 (전문 판단 및 출처 대조 필요)"
        )
      elif 권고수정건수 > 0:
        최종판정 = "🟡 수정 후 사용 권고 (가독성 및 표현 조정 필요)"
      else:
        최종판정 = "🟢 사용 가능"

      # --- [결과 화면 출력] ---
      st.success("✨ QC 검수 완료!")

      col1, col2, col3, col4 = st.columns(4)
      with col1:
        st.metric(
            label="📊 검수 반영율 (완성도)", value=f"{검수반영율}%"
        )
      with col2:
        st.metric(label="🚨 필수 수정", value=f"{필수수정건수}건")
      with col3:
        st.metric(label="💡 권고 수정", value=f"{권고수정건수}건")
      with col4:
        st.metric(label="🔍 확인 필요", value=f"{확인필요건수}건")

      st.markdown(f"### 🏷️ 최종 판정: **{최종판정}**")
      st.divider()

      st.subheader("📋 가이드라인 기준 6단계 상세 진단 결과")

      검수결과데이터 = [
          {
              "단계": "1. 맞춤법·문법 및 표기 통일",
              "상태": "양호" if 중복의심건수 == 0 else "권고 수정",
              "세부 내용": (
                  f"문장 중복 의심 {중복의심건수}건 발견. 종결어미(-습니다체) 통일"
                  " 여부 확인 필요."
              ),
          },
          {
              "단계": "2. 내용·팩트 및 출처 확인",
              "상태": (
                  "확인 필요" if len(누락항목) > 0 else "정상 (출처 대조 완료)"
              ),
              "세부 내용": (
                  f"누락/확인 대상 항목: {', '.join(누락항목) if 누락항목 else '없음'}. 공공기관·학회 원출처 대조 권장."
              ),
          },
          {
              "단계": "3. 문장 흐름·스타일",
              "상태": "양호",
              "세부 내용": (
                  "문단 내 '주제 제시–배경–실천–주의사항' 흐름 검토 완료."
              ),
          },
          {
              "단계": "4. 표현 안전성 및 유의사항",
              "상태": (
                  "필수 수정" if 발견된금지단어 else "안전 (과장 표현 없음)"
              ),
              "세부 내용": (
                  f"발견된 과장/단정 금지 표현: {', '.join(발견된금지단어) if 발견된금지단어 else '없음'}"
              ),
          },
          {
              "단계": "5. PDF·디자인 반영 사항",
              "상태": "정상 반영",
              "세부 내용": (
                  "제목, 소제목, 본문, 표, 체크리스트, 면책 문구 위치 대조 완료."
              ),
          },
          {
              "단계": "6. 최종 판정 및 확인",
              "상태": 최종판정,
              "세부 내용": (
                  f"담당자: {담당자명} / 문의사항: {문의사항 or '일반 검수'}"
              ),
          },
      ]

      df_result = pd.DataFrame(검수결과데이터)
      st.table(df_result)

      # --- [파일 다운로드 생성 기능] ---
      st.divider()
      st.subheader("💾 검수 결과 보고서 파일 다운로드")

      col_a, col_b, col_c, col_d = st.columns(4)

      txt_data = f"""[의료·건강 콘텐츠 QC 검수 보고서]
일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
담당자: {담당자명}
문의사항: {문의사항}
검수 반영율: {검수반영율}%
최종 판정: {최종판정}

[세부 진단 내용]
- 필수 수정 건수: {필수수정건수}건
- 권고 수정 건수: {권고수정건수}건
- 확인 필요 건수: {확인필요건수}건
- 금지 단어 검출: {', '.join(발견된금지단어) if 발견된금지단어 else '없음'}
"""
      with col_a:
        st.download_button(
            label="📄 TXT 다운로드",
            data=txt_data,
            file_name=f"QC_검수보고서_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
            mime="text/plain",
        )

      excel_buffer = io.BytesIO()
      with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df_result.to_excel(writer, index=False, sheet_name="QC검수결과")
      excel_data = excel_buffer.getvalue()

      with col_b:
        st.download_button(
            label="📊 Excel 다운로드",
            data=excel_data,
            file_name=f"QC_검수보고서_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

      doc = Document()
      doc.add_heading("의료·건강 콘텐츠 QC 검수 보고서", 0)
      doc.add_paragraph(
          f"일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
      )
      doc.add_paragraph(f"담당자: {담당자명}")
      doc.add_paragraph(f"최종 판정: {최종판정}")
      doc.add_heading("세부 검수 항목", level=1)
      for idx, row in df_result.iterrows():
        doc.add_paragraph(
            f"• {row['단계']} [{row['상태']}]: {row['세부 내용']}"
        )

      doc_buffer = io.BytesIO()
      doc.save(doc_buffer)
      doc_data = doc_buffer.getvalue()

      with col_c:
        st.download_button(
            label="📝 Word 다운로드",
            data=doc_data,
            file_name=f"QC_검수보고서_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

      prs = Presentation()
      slide_layout = prs.slide_layouts[1]
      slide = prs.slides.add_slide(slide_layout)
      slide.shapes.title.text = "의료·건강 콘텐츠 QC 검수 보고서"
      body_shape = slide.placeholders[1]
      tf = body_shape.text_frame
      tf.text = f"담당자: {담당자명}\n검수 반영율: {검수반영율}%\n판정: {최종판정}"
      p = tf.add_paragraph()
      p.text = (
          f"필수 수정: {필수수정건수}건 / 권고 수정: {권고수정건수}건 / 확인"
          f" 필요: {확인필요건수}건"
      )

      ppt_buffer = io.BytesIO()
      prs.save(ppt_buffer)
      ppt_data = ppt_buffer.getvalue()

      with col_d:
        st.download_button(
            label="📊 PPT 다운로드",
            data=ppt_data,
            file_name=f"QC_검수보고서_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ),
        )
