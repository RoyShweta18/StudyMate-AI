
import tomllib
from google import genai

with open(".streamlit/secrets.toml", "rb") as f:
    secrets = tomllib.load(f)

client = genai.Client(api_key=secrets["GEMINI_API_KEY"])
response = client.models.generate_content(
    model="gemini-3.8-flash-lite",
    contents="Say hello in one short sentence",
)
print(response.text)