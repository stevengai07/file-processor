import re
code = open('app_streamlit.py', encoding='utf-8').read()
code = re.sub(r'st\.markdown\(""".*?""", unsafe_allow_html=True\)', '', code, flags=re.DOTALL, count=1)
open('app_streamlit.py', 'w', encoding='utf-8').write(code)
print('Done - CSS block removed')
