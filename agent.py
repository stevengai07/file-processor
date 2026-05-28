import os
from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from extractor import extract_text
from schema import ExtractedInfo

load_dotenv()

def build_agent():
    llm = ChatOpenAI(model="gpt-4o", temperature=0,
                     openai_api_key=os.getenv("OPENAI_API_KEY"))
    return llm.with_structured_output(ExtractedInfo)

def create_prompt(document_text: str):
    prompt = ChatPromptTemplate.from_messages([
        ("system", "You are an expert document analyst AI agent. Carefully read the document and extract structured key information. Only extract information explicitly present in the text."),
        ("human", "Analyze the following document:\n\n---DOCUMENT START---\n{document}\n---DOCUMENT END---\n\nExtract: title, summary, key topics, named entities, dates, action items, and overall sentiment.")
    ])
    return prompt.format_messages(document=document_text)

def run_agent(file_path: str, ocr_preset: str = "scanner", lang: str = "eng") -> ExtractedInfo:
    print(f"Loading: {file_path}")
    text = extract_text(file_path, ocr_preset=ocr_preset, lang=lang)
    if not text:
        raise ValueError("No text could be extracted from the document.")
    words = text.split()
    if len(words) > 12000:
        text = " ".join(words[:12000])
    print("Running extraction agent...")
    return build_agent().invoke(create_prompt(text))

def display_results(result: ExtractedInfo):
    print("\n" + "="*52)
    print("EXTRACTED INFORMATION")
    print("="*52)
    print(f"Title:     {result.title}")
    print(f"\nSummary:\n  {result.summary}")
    print(f"\nTopics:")
    for t in result.key_topics: print(f"  - {t}")
    print(f"\nEntities:  {', '.join(result.entities)}")
    print(f"\nDates:     {', '.join(result.dates) if result.dates else 'None'}")
    if result.action_items:
        print(f"\nActions:")
        for a in result.action_items: print(f"  - {a}")
    print(f"\nSentiment: {result.sentiment}")
    print("="*52)

if __name__ == "__main__":
    import sys
    file     = sys.argv[1] if len(sys.argv) > 1 else "sample.pdf"
    preset   = sys.argv[2] if len(sys.argv) > 2 else "scanner"
    language = sys.argv[3] if len(sys.argv) > 3 else "eng"
    result   = run_agent(file, ocr_preset=preset, lang=language)
    display_results(result)