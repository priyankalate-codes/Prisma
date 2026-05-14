import os
import zipfile
import json
import tempfile
import logging

# Official Adobe v4+ Imports
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.params.extract_pdf.extract_pdf_params import ExtractPDFParams
from adobe.pdfservices.operation.pdfjobs.params.extract_pdf.extract_element_type import ExtractElementType
from adobe.pdfservices.operation.pdfjobs.jobs.extract_pdf_job import ExtractPDFJob
from adobe.pdfservices.operation.pdfjobs.result.extract_pdf_result import ExtractPDFResult

logger = logging.getLogger(__name__)

def adobe_pdf_extract(pdf_path, client_id, client_secret):
    """
    Extracts semantic JSON using the official Adobe PDF Services SDK (v4+).
    Bypasses manual token handling to prevent 'access_token' errors.
    """
    logger.info("[Adobe SDK] Starting extraction for: %s", pdf_path)
    temp_zip_path = None
    
    try:
        # 1. Setup credentials using your .env keys
        credentials = ServicePrincipalCredentials(
            client_id=client_id,
            client_secret=client_secret
        )

        # 2. Create the main PDF Services instance
        pdf_services = PDFServices(credentials=credentials)

        # 3. Read the input file into memory
        with open(pdf_path, "rb") as file:
            input_stream = file.read()

        # 4. Upload the asset to Adobe
        input_asset = pdf_services.upload(
            input_stream=input_stream, 
            mime_type=PDFServicesMediaType.PDF
        )

        # 5. Build options (Extracting Text and Tables)
        extract_pdf_params = ExtractPDFParams(
            elements_to_extract=[ExtractElementType.TEXT, ExtractElementType.TABLES]
        )

        # 6. Create the Extract Job
        extract_pdf_job = ExtractPDFJob(
            input_asset=input_asset, 
            extract_pdf_params=extract_pdf_params
        )

        # 7. Submit the job and let the SDK handle the polling automatically
        logger.info("[Adobe SDK] Uploading and processing (this takes a few seconds)...")
        location = pdf_services.submit(extract_pdf_job)
        pdf_services_response = pdf_services.get_job_result(location, ExtractPDFResult)

        # 8. Download the resulting ZIP file
        result_asset = pdf_services_response.get_result().get_resource()
        stream_asset = pdf_services.get_content(result_asset)

        with tempfile.NamedTemporaryFile(delete=False, suffix="_adobe_extract.zip") as file:
            temp_zip_path = file.name
            file.write(stream_asset.get_input_stream())

        # 9. Unzip in memory and read the structured JSON
        with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
            with zip_ref.open('structuredData.json') as json_file:
                data = json.load(json_file)

        # 10. Cleanup the temporary ZIP file
        if os.path.exists(temp_zip_path):
            os.remove(temp_zip_path)

        logger.info("[Adobe SDK] Extraction successful")
        return data['elements']

    except Exception as e:
        logger.error("[Adobe SDK Error] %s", str(e))
        raise Exception(f"Official Adobe SDK Extraction Failed: {str(e)}")
    finally:
        if temp_zip_path and os.path.exists(temp_zip_path):
            os.remove(temp_zip_path)
