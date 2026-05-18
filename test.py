import requests

headers = {
    "x-api-key": "YOUR_KEY",
    "anthropic-version": "2023-06-01"
}

r = requests.post(
    "https://api.anthropic.com/v1/messages",
    headers=headers,
    json={
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 10,
        "messages": [{"role": "user", "content": "Hi"}]
    },
    timeout=60
)

print(r.status_code)
print(r.text)