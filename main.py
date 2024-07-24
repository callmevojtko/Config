"""
This code sample shows Prebuilt Read operations with the Azure Form Recognizer client library. 
The async versions of the samples require Python 3.6 or later.

To learn more, please visit the documentation - Quickstart: Document Intelligence (formerly Form Recognizer) SDKs
https://learn.microsoft.com/azure/ai-services/document-intelligence/quickstarts/get-started-sdks-rest-api?pivots=programming-language-python
"""

from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest, ContentFormat

"""
Remember to remove the key from your code when you're done, and never post it publicly. For production, use
secure methods to store and access your credentials. For more information, see 
https://docs.microsoft.com/en-us/azure/cognitive-services/cognitive-services-security?tabs=command-line%2Ccsharp#environment-variables-and-application-configuration
"""
endpoint = "https://config-sheets-auto.cognitiveservices.azure.com/"
key = "0c4138dd9c484b2b985fdfb529522c72"

def format_bounding_box(bounding_box):
    if not bounding_box:
        return "N/A"
    # The new structure is likely a list of floats [x1, y1, x2, y2, ...]
    # Let's format it as pairs of coordinates
    coords = [f"[{bounding_box[i]}, {bounding_box[i+1]}]" for i in range(0, len(bounding_box), 2)]
    return ", ".join(coords)

def analyze_read():
    # sample document
    formUrl = "https://raw.githubusercontent.com/Azure-Samples/cognitive-services-REST-api-samples/master/curl/form-recognizer/sample-layout.pdf"

    document_analysis_client = DocumentIntelligenceClient(
        endpoint=endpoint, credential=AzureKeyCredential(key)
    )
    
    poller = document_analysis_client.begin_analyze_document(
        "prebuilt-read", 
        AnalyzeDocumentRequest(url_source=formUrl),
        output_content_format=ContentFormat.MARKDOWN
    )
    result = poller.result()

    print("Document contains content: ", result.content)
    
    for page in result.pages:
        print(f"----Analyzing Read from page #{page.page_number}----")
        print(f"Page has width: {page.width} and height: {page.height}, measured with unit: {page.unit}")

        for line in page.lines:
            print(f"...Line '{line.content}' within bounding box '{format_bounding_box(line.polygon)}'")

        for word in page.words:
            print(f"...Word '{word.content}' has a confidence of {word.confidence}")

    print("----------------------------------------")


if __name__ == "__main__":
    analyze_read()