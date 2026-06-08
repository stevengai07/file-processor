<<<<<<< HEAD
import os
from dotenv import load_dotenv
from langchain_core.prompts import ChatPromptTemplate
from schema import ExtractedInfo
from extractor import extract_text

load_dotenv()

# ── All Available Models ──────────────────────────────────────────────────────
AVAILABLE_MODELS = {
    # OpenAI
    "gpt-4o":            ("OpenAI",    "~$0.004/doc",   "Most accurate, best for complex docs"),
    "gpt-4o-mini":       ("OpenAI",    "~$0.0002/doc",  "Fast & cheap, great for most docs"),
    "gpt-4-turbo":       ("OpenAI",    "~$0.003/doc",   "Large context, high accuracy"),
    "gpt-3.5-turbo":     ("OpenAI",    "~$0.00005/doc", "Fastest & cheapest"),
    # Anthropic Claude
    "claude-opus-4-5":   ("Anthropic", "~$0.01/doc",    "Most powerful Claude model"),
    "claude-sonnet-4-5": ("Anthropic", "~$0.002/doc",   "Best balance of speed & quality"),
    "claude-haiku-3-5":  ("Anthropic", "~$0.0002/doc",  "Fastest Claude, very affordable"),
    # Google Gemini
    "gemini-2.0-flash":  ("Google",    "~$0.0001/doc",  "Fast, affordable Gemini"),
    "gemini-2.0-pro":    ("Google",    "~$0.003/doc",   "Most capable Gemini model"),
    "gemini-1.5-flash":  ("Google",    "~$0.00005/doc", "Ultra cheap, good quality"),
    # DeepSeek
    "deepseek-chat":     ("DeepSeek",  "~$0.00005/doc", "Excellent quality, very cheap"),
    "deepseek-reasoner": ("DeepSeek",  "~$0.0004/doc",  "Best for complex reasoning tasks"),
    # Grok (xAI)
    "grok-3":            ("xAI",       "~$0.003/doc",   "Most capable Grok model"),
    "grok-3-mini":       ("xAI",       "~$0.0003/doc",  "Fast and affordable Grok"),
    # Qwen (Alibaba)
    "qwen-turbo":        ("Alibaba",   "~$0.00003/doc", "Ultra cheap Qwen model"),
    "qwen-plus":         ("Alibaba",   "~$0.0002/doc",  "Balanced Qwen model"),
    "qwen-max":          ("Alibaba",   "~$0.002/doc",   "Most capable Qwen model"),
}

# ── Provider → ENV key mapping ────────────────────────────────────────────────
PROVIDER_ENV_KEYS = {
    "OpenAI":    "OPENAI_API_KEY",
    "Anthropic": "ANTHROPIC_API_KEY",
    "Google":    "GOOGLE_API_KEY",
    "DeepSeek":  "DEEPSEEK_API_KEY",
    "xAI":       "XAI_API_KEY",
    "Alibaba":   "DASHSCOPE_API_KEY",
}

def get_provider(model: str) -> str:
    return AVAILABLE_MODELS[model][0]

def build_agent(model: str = "gpt-4o-mini"):
    provider = get_provider(model)
    api_key_name = PROVIDER_ENV_KEYS[provider]
    api_key = os.getenv(api_key_name)

    if not api_key:
        raise ValueError(
            f"Missing API key for {provider}. "
            f"Add '{api_key_name}' in the sidebar under API Keys."
        )

    if provider == "OpenAI":
        from langchain_openai import ChatOpenAI
        llm = ChatOpenAI(model=model, temperature=0, openai_api_key=api_key)

    elif provider == "Anthropic":
        from langchain_anthropic import ChatAnthropic
        llm = ChatAnthropic(model=model, temperature=0, anthropic_api_key=api_key)

    elif provider == "Google":
        from langchain_google_genai import ChatGoogleGenerativeAI
        llm = ChatGoogleGenerativeAI(model=model, temperature=0, google_api_key=api_key)

    elif provider == "DeepSeek":
        from langchain_openai import ChatOpenAI
        llm = ChatOpenAI(
            model=model, temperature=0,
            openai_api_key=api_key,
            base_url="https://api.deepseek.com/v1"
        )

    elif provider == "xAI":
        from langchain_openai import ChatOpenAI
        llm = ChatOpenAI(
            model=model, temperature=0,
            openai_api_key=api_key,
            base_url="https://api.x.ai/v1"
        )

    elif provider == "Alibaba":
        from langchain_openai import ChatOpenAI
        llm = ChatOpenAI(
            model=model, temperature=0,
            openai_api_key=api_key,
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
        )

    else:
        raise ValueError(f"Unknown provider: {provider}")

    return llm.with_structured_output(ExtractedInfo)

def create_prompt(document_text: str):
    prompt = ChatPromptTemplate.from_messages([
        ("system", "You are an expert document analyst AI agent. Carefully read the document and extract structured key information. Only extract information explicitly present in the text."),
        ("human", "Analyze the following document:\n\n---DOCUMENT START---\n{document}\n---DOCUMENT END---\n\nExtract: title, summary, key topics, named entities, dates, action items, and overall sentiment.")
    ])
    return prompt.format_messages(document=document_text)

def run_agent(file_path: str, ocr_preset: str = "scanner",
              lang: str = "eng", model: str = "gpt-4o-mini") -> ExtractedInfo:
    print(f"Loading: {file_path}")
    text = extract_text(file_path, ocr_preset=ocr_preset, lang=lang)
    if not text:
        raise ValueError("No text could be extracted from the document.")
    words = text.split()
    if len(words) > 12000:
        text = " ".join(words[:12000])
    print(f"Running extraction agent [{model}]...")
    return build_agent(model=model).invoke(create_prompt(text))

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
    model    = sys.argv[4] if len(sys.argv) > 4 else "gpt-4o-mini"
    result   = run_agent(file, ocr_preset=preset, lang=language, model=model)
 
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
>>>>>>> 834fa95c21c2345785187015bb71077d8712191b
    display_results(result)