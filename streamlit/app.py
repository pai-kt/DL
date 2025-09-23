import streamlit as st
import pandas as pd
import numpy as np

# 제목
st.title("📊 Streamlit 기본 예제")

# 텍스트 출력
st.write("안녕하세요! 이것은 Streamlit 데모 앱입니다.")

# 데이터프레임 예시
df = pd.DataFrame(
    np.random.randn(10, 3),
    columns=["컬럼 A", "컬럼 B", "컬럼 C"]
)
st.dataframe(df)

# 차트 예시
st.line_chart(df)

# 사용자 입력
name = st.text_input("이름을 입력하세요:")
if name:
    st.success(f"환영합니다, {name}님!")

# 버튼 예시
if st.button("버튼 클릭"):
    st.write("버튼이 눌렸습니다 🚀")