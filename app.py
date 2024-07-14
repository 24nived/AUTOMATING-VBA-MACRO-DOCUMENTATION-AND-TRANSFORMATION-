import streamlit as st
from vba_analysis import extract_vba_code_from_excel, parse_vba_code, enhance_function_descriptions
from pyngrok import ngrok

# Set the authtoken for ngrok
ngrok.set_auth_token("2jEPPeMbMW0r5DNj5CthL0pE2j7_2kjjzMahCvqqZBhRtwP9c")

def main():
    st.title("VBA Macro Documentation and Analysis")
    st.write("""
    Upload an Excel file containing VBA macros to extract and analyze the code.
    The tool will provide details about the variables, functions, and comments found in the macros.
    """)

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file:
        vba_code = extract_vba_code_from_excel(uploaded_file)
        variables, functions, comments = parse_vba_code(vba_code)

        st.header("Extracted VBA Code")
        st.text_area("VBA Code", vba_code, height=300)

        st.header("Parsed Information")
        st.subheader("Variables")
        st.write(variables)

        st.subheader("Functions")
        st.write(functions)

        st.subheader("Comments")
        st.write(comments)

        st.subheader("Function Descriptions (Enhanced with spaCy)")
        enhanced_functions = enhance_function_descriptions(functions)
        st.write(enhanced_functions)

if __name__ == "__main__":
    main()

    # Setup a tunnel to the streamlit port 8501
    public_url = ngrok.connect(addr='8501', proto='http')
    print('Streamlit app can be accessed at:', public_url)
