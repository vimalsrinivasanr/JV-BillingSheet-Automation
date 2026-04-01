import google.generativeai as genai
import pandas as pd
import json

class AIMapper:
    """Uses Gemini to intelligently map raw spreadsheet columns to target business fields."""

    def __init__(self, api_key):
        if api_key:
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel('gemini-1.5-flash')
        else:
            self.model = None

    def analyze_template(self, df_sample):
        """
        Analyzes the first 5 rows of a spreadsheet to determine column mappings.
        Returns a dictionary mapping internal keys to Excel column indices.
        """
        if not self.model:
            return None # Fallback to hardcoded logic

        # Prepare the sample data for the prompt
        headers = df_sample.columns.tolist()
        sample_data = df_sample.head(3).to_dict(orient='records')
        
        prompt = f"""
        Act as a Financial Data Engineer. I have a billing spreadsheet with these headers: {headers}
        And this sample data: {json.dumps(sample_data)}
        
        I need to map these columns to our internal SAP processing fields:
        - workday_id (The unique employee identifier)
        - cap_center (The department or capability center name)
        - legal_entity (The company or entity name)
        - classification (Whether the row is 'Billable', 'Non Billable', etc.)
        - billed_status (Whether it is 'Billed' or 'Unbilled')
        - invoice_no (The document or invoice number)
        - ic_code (The entity code like NL_RSH, US_SAP_L)
        
        Return ONLY a JSON object mapping each field to its 0-indexed column position.
        Example: {{"workday_id": 0, "cap_center": 6, ...}}
        """
        
        try:
            response = self.model.generate_content(prompt)
            # Extracted JSON from response
            text = response.text.strip()
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0].strip()
            return json.loads(text)
        except Exception as e:
            print(f"AI Mapper Error: {e}")
            return None
