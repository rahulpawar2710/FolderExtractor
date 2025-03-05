from openai import OpenAI
from pathlib import Path
import pandas as pd
import json
import re
import time
import os
from openpyxl import load_workbook
from datetime import datetime
from lxml import etree
import streamlit as st
import tempfile

# Streamlit App
st.title("Invoice Data Extractor")

# Load OpenAI API key from Streamlit Secrets
if "OPENAI_API_KEY" not in st.secrets:
    st.error("❌ OpenAI API key is missing! Please add it in Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


# Function to parse XML content
def parse_xml(xml_content):
    return etree.XML(xml_content)


# Function to get element by XPath
def get_element_by_full_xpath(tree, xpath_expr):
    elements = tree.xpath(xpath_expr)
    return elements[0] if elements else None


# Function to process invoices
def process_invoices(uploaded_files):
    if not uploaded_files:
        st.error("⚠️ No files uploaded! Please upload at least one XML and one PDF.")
        return None

    # Read Config.xml from uploaded files
    config_file = next((f for f in uploaded_files if f.name.lower() == "config.xml"), None)
    if not config_file:
        st.error("⚠️ Config.xml file is required! Please upload it along with PDFs.")
        return None

    xml_content = config_file.read()
    tree = parse_xml(xml_content)

    # Get OpenAI Configuration from XML
    try:
        ActiveEnvironmentPath = '/Configuration/ActiveEnvironment'
        ActiveEnvironmentValue = get_element_by_full_xpath(tree, ActiveEnvironmentPath).text
        ContextualOpenAiModelName = get_element_by_full_xpath(tree, f"/Configuration/{ActiveEnvironmentValue}/ContextualOpenAiModelName").text
        VectorName = get_element_by_full_xpath(tree, f"/Configuration/{ActiveEnvironmentValue}/VectorName").text
    except AttributeError:
        st.error("⚠️ Error reading Config.xml. Make sure the XML structure is correct.")
        return None

    # Initialize OpenAI Assistant
    try:
        assistant = client.beta.assistants.create(
            name="Invoice Processing Assistant",
            instructions="You are an expert invoice analyst. Extract key invoice details accurately.",
            model=ContextualOpenAiModelName,
            tools=[{"type": "file_search"}],
        )

        # Create vector store
        vector_store = client.beta.vector_stores.create(name=VectorName)
        vector_store_id = vector_store.id
    except Exception as e:
        st.error(f"❌ Error initializing OpenAI Assistant: {e}")
        return None

    output_data = []

    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".pdf"):
            st.write(f"Processing: {uploaded_file.name}")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
                temp_file.write(uploaded_file.read())
                temp_file_path = temp_file.name

            try:
                # Upload file to OpenAI Vector Store
                with open(temp_file_path, "rb") as f:
                    file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
                        vector_store_id=vector_store_id, files=[f]
                    )
                    st.write(f"✅ File {uploaded_file.name} uploaded.")

                # Extract invoice details
                thread = client.beta.threads.create(
                    messages=[
                        {
                            "role": "user",
                            "content": (
                                "Please extract the following details in JSON format: "
                                "Invoice Number, Invoice Date, Purchase Order, Subtotal, HST, Total."
                            )
                        }
                    ]
                )

                run = client.beta.threads.runs.create_and_poll(
                    thread_id=thread.id, assistant_id=assistant.id
                )

                messages = list(client.beta.threads.messages.list(thread_id=thread.id, run_id=run.id))

                # Ensure messages are not empty
                if not messages:
                    st.warning(f"⚠️ No response received for {uploaded_file.name}. Skipping...")
                    continue

                message_content_obj = messages[0].content[0]

                # Ensure message_content is a string
                if hasattr(message_content_obj, "text"):
                    message_content = message_content_obj.text
                else:
                    st.warning(f"⚠️ No valid text response for {uploaded_file.name}. Skipping...")
                    continue

                # Extract JSON data using regex
                match = re.search(r"\{.*\}", message_content, re.DOTALL)

                if match:
                    json_data = match.group(0)
                    try:
                        parsed_result = json.loads(json_data)
                        parsed_result["FileName"] = uploaded_file.name
                        output_data.append(parsed_result)
                    except json.JSONDecodeError:
                        st.error(f"❌ Error parsing JSON for {uploaded_file.name}")

            except Exception as e:
                st.error(f"❌ Error processing {uploaded_file.name}: {e}")

    if output_data:
        df = pd.DataFrame(output_data)
        st.dataframe(df)

        # Save output file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as output_file:
            df.to_excel(output_file.name, index=False)
            output_path = output_file.name

        st.download_button(
            label="📥 Download Extracted Data",
            data=open(output_path, "rb").read(),
            file_name="Extracted_Invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# File Upload Section
uploaded_files = st.file_uploader("📂 Upload invoice PDFs & Config.xml", type=["pdf", "xml"], accept_multiple_files=True)

if st.button("🔍 Extract Data"):
    process_invoices(uploaded_files)
