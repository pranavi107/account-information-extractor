import streamlit as st
import pandas as pd
import os
import zipfile

from extractor import process_document

st.set_page_config(layout="wide")

st.title("Account Information Extractor")

# ---------------------------------------------------
# UPLOAD WORD DOCUMENTS
# ---------------------------------------------------

uploaded_docs = st.file_uploader(
    "Upload Word Documents",
    type=["docx"],
    accept_multiple_files=True
)

# ---------------------------------------------------
# UPLOAD EXCEL FILE
# ---------------------------------------------------

uploaded_excel = st.file_uploader(
    "Upload Excel File with Account Numbers",
    type=["xlsx"]
)

# ---------------------------------------------------
# BUTTON
# ---------------------------------------------------

if st.button("Generate All Documents"):

    # ---------------------------------------------------
    # VALIDATION
    # ---------------------------------------------------

    if not uploaded_docs:

        st.error(
            "Please upload Word documents."
        )

    elif not uploaded_excel:

        st.error(
            "Please upload Excel file."
        )

    else:

        try:

            # ---------------------------------------------------
            # CREATE OUTPUT FOLDER
            # ---------------------------------------------------

            output_folder = "outputs"

            os.makedirs(
                output_folder,
                exist_ok=True
            )

            # ---------------------------------------------------
            # CLEAR OLD FILES
            # ---------------------------------------------------

            for file_name in os.listdir(output_folder):

                file_path = os.path.join(
                    output_folder,
                    file_name
                )

                if os.path.isfile(file_path):

                    os.remove(file_path)

            # ---------------------------------------------------
            # READ EXCEL
            # ---------------------------------------------------

            df = pd.read_excel(uploaded_excel)

            # ---------------------------------------------------
            # REMOVE DUPLICATES
            # ---------------------------------------------------

            account_numbers = (
                df.iloc[:, 0]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )

            total_accounts = len(account_numbers)

            total_generated = 0

            # ---------------------------------------------------
            # PROCESS EACH ACCOUNT
            # ---------------------------------------------------

            for account_number in account_numbers:

                found = False

                for doc_file in uploaded_docs:

                    try:

                        output_path = process_document(
                            doc_file,
                            account_number
                        )

                        if output_path:

                            found = True

                            total_generated += 1

                            st.success(
                                f"Generated: {account_number}.docx"
                            )

                            break

                    except Exception as e:

                        st.error(
                            f"{account_number}: {str(e)}"
                        )

                if not found:

                    st.warning(
                        f"Account not found: {account_number}"
                    )

            # ---------------------------------------------------
            # CREATE ZIP FILE
            # ---------------------------------------------------

            zip_path = "outputs.zip"

            with zipfile.ZipFile(
                zip_path,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zipf:

                for file_name in os.listdir(output_folder):

                    file_path = os.path.join(
                        output_folder,
                        file_name
                    )

                    zipf.write(
                        file_path,
                        arcname=file_name
                    )

            # ---------------------------------------------------
            # SUCCESS MESSAGE
            # ---------------------------------------------------

            st.success(
                f"""
Completed Successfully

Unique Accounts Processed: {total_accounts}

Files Generated: {total_generated}
"""
            )

            # ---------------------------------------------------
            # DOWNLOAD ZIP
            # ---------------------------------------------------

            with open(zip_path, "rb") as f:

                zip_bytes = f.read()

            st.download_button(
                label="Download All Documents ZIP",
                data=zip_bytes,
                file_name="outputs.zip",
                mime="application/zip"
            )

        except Exception as e:

            st.error(str(e))